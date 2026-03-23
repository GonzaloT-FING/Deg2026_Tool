import os
import numpy as np
import pandas as pd
import torch
import functorch
import matplotlib as mpl
import matplotlib.pyplot as plt
import logging
import collections
from typing import Callable, Optional, Dict, Union, Sequence, Tuple
from torchmin import minimize, ScipyMinimizer, least_squares
from datetime import datetime

logger = logging.getLogger(__name__)
plt.style.use("default")
mpl.rcParams["figure.facecolor"] = "white"
plt.rcParams["font.weight"] = "bold"
plt.rcParams["axes.linewidth"] = 1.5
mpl.rcParams["ps.fonttype"] = 42


class Multieis:
    """
    An immittance batch processing class

    :param p0: A 1D or 2D tensor of initial guess

    :param freq: An (m, ) 1D tensor containing the frequencies. \
                 Where m is the number of frequencies

    :param Z: An (m, n) 2D tensor of complex immittances. \
              Where m is the number of frequencies and \
              n is the number of spectra

    :param bounds: A sequence of (min, max) pairs for \
                   each element in p0. The values must be real

    :param smf: A tensor of real elements same size as p0. \
                when set to inf, the corresponding parameter is kept constant

    :param func: A model e.g an equivalent circuit model (ECM) or \
                 an arbitrary immittance expression composed as python function

    :param weight: A string representing the weighting scheme or \
                   an (m,n) 2-D tensor of real values containing \
                   the measurement standard deviation. \
                   Defaults to unit weighting if left unspecified.

    :param immittance: A string corresponding to the immittance type

    :returns: A Multieis instance

    """

    def __init__(
        self,
        p0: torch.tensor,
        freq: torch.tensor,
        Z: torch.tensor,
        bounds: Sequence[Union[int, float]],
        smf: torch.tensor,
        func: Callable[[float, float], float],
        immittance: str = "impedance",
        weight: Optional[Union[str, torch.tensor]] = None,
    ) -> None:

        assert (
            p0.ndim > 0 and p0.ndim <= 2
        ), ("Initial guess must be a 1-D tensor or 2-D "
            "tensor with same number of cols as `F`")
        assert (
            Z.ndim == 2 and Z.shape[1] >= 5
        ), "The algorithm requires that the number of spectra be >= 5"
        assert freq.ndim == 1, "The frequencies supplied should be 1-D"
        assert (
            len(freq) == Z.shape[0]
        ), ("Length mismatch: The len of F is {} while the rows of Z are {}"
            .format(len(freq), Z.shape[0]))

        # Create the lower and upper bounds
        try:
            self.lb = self.check_zero_and_negative_values(
                torch.as_tensor([i[0] for i in bounds])
            )
            self.ub = self.check_zero_and_negative_values(
                torch.as_tensor([i[1] for i in bounds])
            )
        except IndexError:
            print("Bounds must be a sequence of min-max pairs")

        if p0.ndim == 1:
            self.p0 = self.check_zero_and_negative_values(self.check_nan_values(p0))
            self.num_params = len(self.p0)
            assert (
                len(self.lb) == self.num_params
            ), "Shape mismatch between initial guess and bounds"
            if __debug__:
                if not torch.all(
                    torch.logical_and(
                        torch.greater(self.p0, self.lb),
                        torch.greater(self.ub, self.p0)
                    )
                ):
                    raise AssertionError("""Initial guess can not be
                                        greater than the upper bound
                                        or less than lower bound""")
        elif (p0.ndim == 2) and (1 in p0.shape):
            self.p0 = self.check_zero_and_negative_values(self.check_nan_values(p0.flatten()))
            self.num_params = len(self.p0)
            assert (
                len(self.lb) == self.num_params
            ), "Shape mismatch between initial guess and bounds"
            if __debug__:
                if not torch.all(
                    torch.logical_and(
                        torch.greater(self.p0, self.lb),
                        torch.greater(self.ub, self.p0)
                    )
                ):
                    raise AssertionError("""Initial guess can not be
                                        greater than the upper bound
                                        or less than lower bound""")
        else:
            assert p0.shape[1] == Z.shape[1], ("Columns of p0 "
                                               "do not match that of Z")
            assert (
                len(self.lb) == p0.shape[0]
            ), ("The len of p0 is {} while that of the bounds is {}"
                .format(p0.shape[0], len(self.lb)))
            self.p0 = self.check_zero_and_negative_values(self.check_nan_values(p0))
            self.num_params = p0.shape[0]

        self.immittance_list = ["admittance", "impedance"]
        assert (
            immittance.lower() in self.immittance_list
        ), "Either use 'admittance' or 'impedance'"

        self.num_freq = len(freq)
        self.num_eis = Z.shape[1]
        self.F = freq.type(torch.FloatTensor)
        self.Z = self.check_is_complex(Z)
        self.Z_exp = self.Z.clone()
        self.Y_exp = 1 / self.Z_exp.clone()
        self.indices = None
        self.n_fits = None

        self.func = func
        self.immittance = immittance

        self.smf = smf.type(torch.FloatTensor)

        self.smf_1 = torch.where(torch.isinf(self.smf), 0.0, self.smf)

        self.kvals = torch.cumsum(
            np.insert(torch.where(
                torch.isinf(self.smf), 1, self.num_eis), 0, 0), dim=0
        )

        self.gather_indices = torch.zeros(size=(self.num_params, self.num_eis), dtype=torch.int64)
        for i in range(self.num_params):
            self.gather_indices[i, :] = torch.arange(self.kvals[i], self.kvals[i+1])

        self.d2m = self.get_fd()
        self.dof = (2 * self.num_freq * self.num_eis) - \
            (self.num_params * self.num_eis)

        self.plot_title1 = " ".join(
            [x.title() for x in self.immittance_list if (x == self.immittance)]
        )
        self.plot_title2 = " ".join(
            [x.title() for x in self.immittance_list if x != self.immittance]
        )

        self.lb_vec, self.ub_vec = self.get_bounds_vector(self.lb, self.ub)

        # Define weighting strategies
        if torch.is_tensor(weight):
            self.weight_name = "sigma"
            assert (
                Z.shape == weight.shape
            ), "Shape mismatch between Z and the weight tensor"
            self.Zerr_Re = weight
            self.Zerr_Im = weight
        elif isinstance(weight, str):
            assert weight.lower() in [
                "proportional",
                "modulus",
            ], ("weight must be one of None, "
                "proportional', 'modulus' or an 2-D tensor of weights")
            if weight.lower() == "proportional":
                self.weight_name = "proportional"
                self.Zerr_Re = self.Z.real
                self.Zerr_Im = self.Z.imag
            else:
                self.weight_name = "modulus"
                self.Zerr_Re = torch.abs(self.Z)
                self.Zerr_Im = torch.abs(self.Z)
        elif weight is None:
            # if set to None we use "unit" weighting
            self.weight_name = "unit"
            self.Zerr_Re = torch.ones(self.num_freq, self.num_eis)
            self.Zerr_Im = torch.ones(self.num_freq, self.num_eis)
        else:
            raise AttributeError(
                ("weight must be one of None, "
                 "proportional', 'modulus' or an 2-D tensor of weights")
            )

    def __str__(self):
        return f"""Multieis({self.p0},{self.F},{self.Z},{self.Zerr_Re},\
                {self.Zerr_Im}, {list(zip(self.lb, self.ub))},\
                {self.func},{self.immittance},{self.weight_name})"""

    __repr__ = __str__

    @staticmethod
    def check_nan_values(arr):
        if torch.isnan(torch.sum(arr)):
            raise Exception("Values must not contain nan")
        else:
            return arr

    @staticmethod
    def check_zero_and_negative_values(arr):
        if torch.all(arr > 0):
            return arr
        raise Exception("Values must be greater than zero")

    @staticmethod
    def try_convert(x):
        try:
            return str(x)
        except Exception as e:
            print(e.__doc__)
            print(e.message)
        return x

    @staticmethod
    def check_is_complex(arr):
        if torch.is_complex(arr):
            return arr.type(torch.complex64)
        else:
            return arr.type(torch.complex64)

    def get_bounds_vector(self,
                          lb: torch.tensor,
                          ub: torch.tensor
                          ) -> Tuple[torch.tensor, torch.tensor]:
        """
        Creates vectors for the upper and lower \
        bounds which are the same length \
        as the number of parameters to be fitted.

        :param lb: A 1D tensor of lower bounds
        :param ub: A 1D tensor of upper bounds

        :returns: A tuple of bounds vectors

        """
        lb_vec = torch.zeros(
            self.num_params * self.num_eis
            - (self.num_eis - 1)
            * torch.sum(torch.isinf(self.smf))
        )
        ub_vec = torch.zeros(
            self.num_params * self.num_eis
            - (self.num_eis - 1) * torch.sum(torch.isinf(self.smf))
        )
        for i in range(self.num_params):
            lb_vec[self.kvals[i]:self.kvals[i + 1]] = lb[i]
            ub_vec[self.kvals[i]:self.kvals[i + 1]] = ub[i]
        return lb_vec, ub_vec

    def get_fd(self):
        """
        Computes the finite difference stencil \
        for a second order derivative. \
        The derivatives at the boundaries is calculated \
        using special finite difference equations
        derived specifically for just these points \
        (aka higher order boundary conditions).
        They are used to handle numerical problems \
        that occur at the edge of grids.

        :returns: Finite difference stencil for a second order derivative
        """
        self.d2m = torch.zeros(size=(self.num_eis, self.num_eis), dtype=torch.float32)
        self.d2m[0, :4] = torch.tensor([2, -5, 4, -1])
        for k in range(1, self.num_eis - 1):
            self.d2m[k, k - 1:k + 2] = torch.tensor([1, -2, 1])
        self.d2m[-1, -4:] = torch.tensor([-1, 4, -5, 2])
        return self.d2m

    def convert_to_internal(self,
                            p: torch.tensor
                            ) -> torch.tensor:
        """
        Converts A tensor of parameters from an external \
        to an internal coordinates (log10 scale)

        :param p: A 1D or 2D tensor of parameter values

        :returns: Parameters in log10 scale
        """
        assert p.ndim > 0 and p.ndim <= 2
        if p.ndim == 1:
            par = torch.broadcast_to(p[:, None],
                                     (self.num_params, self.num_eis))
        else:
            par = p
        self.p0_mat = torch.zeros(
            self.num_params * self.num_eis
            - (self.num_eis - 1) * torch.sum(torch.isinf(self.smf))
        )
        for i in range(self.num_params):
            self.p0_mat[self.kvals[i]:self.kvals[i + 1]] = par[
                i, : self.kvals[i + 1] - self.kvals[i]
            ]
        p_log = torch.log10(
            (self.p0_mat - self.lb_vec) / (1 - self.p0_mat / self.ub_vec)
        )
        return p_log

    def convert_to_external(self, P: torch.tensor) -> torch.tensor:

        """
        Converts A tensor of parameters from an internal \
        to an external coordinates

        :param p: A 1D tensor of parameter values

        :returns: Parameters in normal scale
        """
        par_ext = torch.zeros(self.num_params, self.num_eis)
        for i in range(self.num_params):
            par_ext[i, :] = (
                self.lb_vec[self.kvals[i]:self.kvals[i + 1]]
                + 10 ** P[self.kvals[i]:self.kvals[i + 1]]
            ) / (
                1
                + (10 ** P[self.kvals[i]:self.kvals[i + 1]])
                / self.ub_vec[self.kvals[i]:self.kvals[i + 1]]
            )
        return par_ext

    def compute_wrss(self,
                     p: torch.tensor,
                     f: torch.tensor,
                     z: torch.tensor,
                     zerr_re: torch.tensor,
                     zerr_im: torch.tensor
                     ) -> torch.tensor:

        """
        Computes the scalar weighted residual sum of squares \
        (aka scaled version of the chisquare or the chisquare itself)

        :param p: A 1D tensor of parameter values

        :param f: A 1D tensor of frequency

        :param z: A 1D tensor of complex immittances

        :param zerr_re: A 1D tensor of weights for \
                        the real part of the immittance

        :param zerr_im: A 1D tensor of weights for \
                        the imaginary part of the immittance

        :returns: A scalar value of the \
                  weighted residual sum of squares

        """
        z_concat = torch.cat([z.real, z.imag], dim=0)
        sigma = torch.cat([zerr_re, zerr_im], dim=0)
        z_model = self.func(p, f)
        wrss = torch.linalg.vector_norm(((z_concat - z_model) / sigma)) ** 2
        return wrss

    def compute_rss(self,
                    p: torch.tensor,
                    f: torch.tensor,
                    z: torch.tensor,
                    zerr_re: torch.tensor,
                    zerr_im: torch.tensor,
                    lb,
                    ub
                    ) -> torch.tensor:
        """
        Computes the vector of weighted residuals. \
        This is the objective function passed to the least squares solver.

        :param p: A 1D tensor of parameter values

        :param f: A 1D tensor of frequency

        :param z: A 1D tensor of complex immittances

        :param zerr_re: A 1D tensor of weights for \
                        the real part of the immittance

        :param zerr_im: A 1D tensor of weights for \
                        the imaginary part of the immittance

        :param lb: A 1D tensor of values for the lower bounds

        :param ub: A 1D tensor of values for the upper bounds

        :returns: A vector of residuals

        """
        p = (lb + 10 ** (p)) / (1 + 10 ** (p) / ub)
        z_concat = torch.cat([z.real, z.imag], dim=0)
        sigma = torch.cat([zerr_re, zerr_im], dim=0)
        z_model = self.func(p, f)
        residuals = 0.5 * ((z_concat - z_model) / sigma)
        return residuals

    def compute_wrms(self,
                     p: torch.tensor,
                     f: torch.tensor,
                     z: torch.tensor,
                     zerr_re: torch.tensor,
                     zerr_im: torch.tensor
                     ) -> torch.tensor:
        """
        Computes the weighted residual mean square

        :param p: A 1D tensor of parameter values

        :param f: A 1D tensor of frequency

        :param z: A 1D tensor of complex immittance

        :param zerr_re: A 1D tensor of weights for \
                        the real part of the immittance

        :param zerr_im: A 1D tensor of weights for \
                        the imaginary part of the immittance

        :returns: A scalar value of the weighted residual mean square
        """
        z_concat = torch.cat([z.real, z.imag], dim=0)
        sigma = torch.cat([zerr_re, zerr_im], dim=0)
        z_model = self.func(p, f)
        wrss = torch.linalg.vector_norm(((z_concat - z_model) / sigma)) ** 2
        wrms = wrss / (2 * len(f) - len(p))
        return wrms

    def compute_perr(self,
                     P: torch.tensor,
                     F: torch.tensor,
                     Z: torch.tensor,
                     Zerr_Re: torch.tensor,
                     Zerr_Im: torch.tensor,
                     LB: torch.tensor,
                     UB: torch.tensor,
                     smf: torch.tensor
                     ) -> torch.tensor:

        """
        Computes the error on the parameters resulting from the batch fit
        using the hessian inverse of the parameters at the minimum computed
        via automatic differentiation

        :param P: A 2D tensor of parameter values

        :param F: A 1D tensor of frequency

        :param Z: A 2D tensor of complex immittances

        :param Zerr_Re: A 2D tensor of weights for \
                        the real part of the immittance

        :param Zerr_Im: A 2D tensor of weights for \
                        the imaginary part of the immittance

        :param LB: A 1D tensor of values for \
                   the lower bounds (for the total parameters)

        :param UB: A 1D tensor of values for \
                   the upper bounds (for the total parameters)

        :param smf: A tensor of real elements same size as p0. \
                    when set to inf, the corresponding parameter is kept constant

        :returns: A 2D tensor of the standard error on the parameters

        """
        P_log = self.convert_to_internal(P)

        chitot = self.compute_total_obj(P_log, F, Z, Zerr_Re, Zerr_Im, LB, UB, smf)/self.dof
        hess_mat = torch.autograd.functional.hessian(
            self.compute_total_obj, (P_log, F, Z, Zerr_Re, Zerr_Im, LB, UB, smf)
        )[0][0]
        try:
            # Here we check to see if the Hessian matrix is singular \
            # or ill-conditioned since this makes accurate computation of the
            # confidence intervals close to impossible.
            hess_inv = torch.linalg.inv(hess_mat)
        except torch.linalg.LinAlgError:
            hess_inv = torch.linalg.pinv(hess_mat)

        # The covariance matrix of the parameter estimates
        # is (asymptotically) the inverse of the hessian matrix
        self.cov_mat = hess_inv * chitot
        std_error = torch.sqrt(torch.diag(self.cov_mat))
        perr = torch.zeros(self.num_params, self.num_eis)
        for i in range(self.num_params):
            perr[i, :] = std_error[self.kvals[i]:self.kvals[i + 1]]

        perr = perr.detach().clone() * P
        return torch.nan_to_num(perr, nan=1.0e15)

    def compute_total_obj(self,
                          P: torch.tensor,
                          F: torch.tensor,
                          Z: torch.tensor,
                          Zerr_Re: torch.tensor,
                          Zerr_Im: torch.tensor,
                          LB: torch.tensor,
                          UB: torch.tensor,
                          smf: torch.tensor
                          ) -> torch.tensor:
        """
        This function computes the total scalar objective function to minimize
        which is a combination of the weighted residual sum of squares
        and the (second derivative of the params + smoothing factor)

        :param P: A 1D tensor of parameter values in \
                  log scale

        :param F: A 1D tensor of frequency

        :param Z: A 2D tensor of complex immittances

        :param Zerr_Re: A 2D tensor of weights for \
                        the real part of the immittance

        :param Zerr_Im: A 2D tensor of weights for \
                        the imaginary part of the immittance

        :param LB: A 1D tensor of values for \
                   the lower bounds (for the total parameters)

        :param UB: A 1D tensor of values for \
                   the upper bounds (for the total parameters)

        :param smf: A tensor of real elements same size as p0. \
            when set to inf, the corresponding parameter is kept constant

        :returns: A scalar value of the total objective function

        """

        P_log = torch.take(P, self.gather_indices)

        up = (10 ** P_log)

        P_norm = (
            torch.take(LB, self.gather_indices)
            + up
        ) / (
            1
            + up
            / torch.take(UB, self.gather_indices)
        )

        chi_smf = ((((self.d2m @ P_log.T.float()) * (self.d2m @ P_log.T.float())))
                   .sum(0) * smf).sum()
        wrss_tot = functorch.vmap(self.compute_wrss, in_dims=(1, None, 1, 1, 1))(
            P_norm, F, Z, Zerr_Re, Zerr_Im
        )
        return (torch.sum(wrss_tot) + chi_smf)

    def compute_perr_QR(self,
                        P: torch.tensor,
                        F: torch.tensor,
                        Z: torch.tensor,
                        Zerr_Re: torch.tensor,
                        Zerr_Im: torch.tensor
                        ) -> torch.tensor:

        """
        Computes the error on the parameters resulting from the batch fit
        using QR decomposition

        :param P: A 2D tensor of parameter values

        :param F: A 1D tensor of frequency

        :param Z: A 2D tensor of complex immittances

        :param Zerr_Re: A 2D tensor of weights for \
                        the real part of the immittance

        :param Zerr_Im: A 2D tensor of weights for \
                        the imaginary part of the immittance

        :returns: A 2D tensor of the standard error on the parameters

        Notes
        ----
        Bates, D. M., Watts, D. G. (1988). Nonlinear regression analysis \
        and its applications. New York [u.a.]: Wiley. ISBN: 0471816434
        """
        def grad_func(p, f):
            return torch.autograd.functional.jacobian(self.func, (p, f))[0]
        perr = torch.zeros(self.num_params, self.num_eis)
        for i in range(self.num_eis):
            wrms = self.compute_wrms(P[:, i], F, Z[:, i], Zerr_Re[:, i], Zerr_Im[:, i])
            gradsre = grad_func(P[:, i], F)[:self.num_freq]
            gradsim = grad_func(P[:, i], F)[self.num_freq:]
            diag_wtre_matrix = torch.diag((1/Zerr_Re[:, i]))
            diag_wtim_matrix = torch.diag((1/Zerr_Im[:, i]))
            vre = diag_wtre_matrix.double()@gradsre.double()
            vim = diag_wtim_matrix.double()@gradsim.double()
            Q1, R1 = torch.linalg.qr(torch.cat([vre , vim], dim=0))
            try:
                # Here we check to see if the Hessian matrix is singular or
                # ill-conditioned since this makes accurate computation of the
                # confidence intervals close to impossible.
                invR1 = torch.linalg.inv(R1)
            except torch.linalg.LinAlgError:
                print(f"\nHessian Matrix is singular for spectra {i}")
                invR1 = torch.linalg.pinv(R1)

            perr[:, i] = torch.linalg.vector_norm(invR1, dim=1)*torch.sqrt(wrms)
        # if the error is nan, a value of 1 is assigned.
        return torch.nan_to_num(perr, nan=1.0e15)

    def compute_aic(self,
                    p: torch.tensor,
                    f: torch.tensor,
                    z: torch.tensor,
                    zerr_re: torch.tensor,
                    zerr_im: torch.tensor,
                    ) -> torch.tensor:
        """
        Computes the Akaike Information Criterion according to
        `M. Ingdal et al <https://www.sciencedirect.com/science/article/abs/pii/S0013468619311739>`_

        :param p: A 1D tensor of parameter values

        :param f: A 1D tensor of frequency

        :param z: A 1D tensor of complex immittances

        :param zerr_re: A 1D tensor of weights for \
                        the real part of the immittance

        :param zerr_im: A 1D tensor of weights for \
                        the imaginary part of the immittance


        :returns: A value for the AIC
        """

        wrss = self.compute_wrss(p, f, z, zerr_re, zerr_im)
        if self.weight_name == "sigma":
            m2lnL = (
                (2 * self.num_freq) * torch.log(torch.tensor(2 * torch.pi))
                + torch.sum(torch.log(zerr_re**2))
                + torch.sum(torch.log(zerr_im**2))
                + (wrss)
            )
            aic = m2lnL + 2 * self.num_params

        elif self.weight_name == "unit":
            m2lnL = (
                2 * self.num_freq * torch.log(torch.tensor(2 * torch.pi))
                - 2 * self.num_freq
                * torch.log(torch.tensor(2 * self.num_freq))
                + 2 * self.num_freq
                + 2 * self.num_freq * torch.log(wrss)
            )
            aic = m2lnL + 2 * self.num_params

        else:
            wt_re = 1 / zerr_re**2
            wt_im = wt_re
            m2lnL = (
                2 * self.num_freq * torch.log(torch.tensor(2 * torch.pi))
                - 2 * self.num_freq
                * torch.log(torch.tensor(2 * self.num_freq))
                + 2 * self.num_freq
                - torch.sum(torch.log(wt_re))
                - torch.sum(torch.log(wt_im))
                + 2 * self.num_freq * torch.log(wrss)
            )  # log-likelihood calculation
            aic = m2lnL + 2 * (self.num_params + 1)
        return aic

    def fit_simultaneous(self,
                         method: str = 'bfgs',
                         n_iter: int = 5000,
                         ) -> Tuple[
        torch.tensor, torch.tensor, torch.tensor, torch.tensor, torch.tensor
    ]:  # Optimal parameters, parameter error,
        # weighted residual mean square, and the AIC

        """
        Simultaneous fitting routine with an arbitrary smoothing factor..

        :params method: Solver to use (must be one of "'TNC', \
                        'BFGS' or 'L-BFGS-B'")

        :params n_iter: Number of iterations

        :returns: A tuple containing the optimal parameters (popt), \
                  the standard error of the parameters (perr), \
                  the objective function at the minimum (chisqr), \
                  the total cost function (chitot) and the AIC
        """
        self.method = method.lower()
        assert (self.method in ['tnc', 'bfgs', 'l-bfgs-b']), ("method must be one of "
                                                              "'TNC', 'BFGS' or 'L-BFGS-B'")

        if hasattr(self, "popt") and self.popt.shape[1] == self.Z.shape[1]:
            print("\nUsing prefit")
            self.par_log = (
                self.convert_to_internal(self.check_nan_values(self.popt)).type(torch.DoubleTensor)
            ).requires_grad_(True)
        else:
            print("\nUsing initial")
            self.par_log = (
                self.convert_to_internal(self.p0).type(torch.DoubleTensor)
            ).requires_grad_(True)

        start = datetime.now()
        optimizer = ScipyMinimizer(
            params=[self.par_log],
            method=self.method,
            tol=1e-16,
            options={"maxfun" if self.method == "tnc" else "maxiter": n_iter}
            )
        self.iteration = 0

        def closure():
            optimizer.zero_grad()
            loss = self.compute_total_obj(
                self.par_log,
                self.F,
                self.Z,
                self.Zerr_Re,
                self.Zerr_Im,
                self.lb_vec,
                self.ub_vec,
                self.smf_1,
            )
            loss.backward()
            print(
                "\rIteration : {}, Loss : {:.5e}"
                .format(self.iteration, loss.detach().clone()/self.dof),
                end="",
            )
            if self.iteration % 1000 == 0:
                print("")
            self.iteration += 1
            return loss

        optimizer.step(closure)

        self.popt = self.convert_to_external(self.par_log).detach()
        self.chitot = torch.as_tensor(optimizer._result.fun)/self.dof

        self.perr = self.compute_perr(
            self.popt,
            self.F,
            self.Z,
            self.Zerr_Re,
            self.Zerr_Im,
            self.lb_vec,
            self.ub_vec,
            self.smf_1,
        )

        self.chisqr = torch.mean(
            functorch.vmap(self.compute_wrms, in_dims=(1, None, 1, 1, 1))(
                self.popt, self.F, self.Z, self.Zerr_Re, self.Zerr_Im
            )
        )
        self.AIC = torch.mean(
            functorch.vmap(self.compute_aic, in_dims=(1, None, 1, 1, 1))(
                self.popt, self.F, self.Z, self.Zerr_Re, self.Zerr_Im
            )
        )
        print("\nOptimization complete")
        end = datetime.now()
        print(f"total time is {end-start}\n", end=" ")
        self.Z_exp = self.Z.clone()
        self.Y_exp = 1 / self.Z_exp.clone()
        self.Z_pred, self.Y_pred = self.model_prediction(self.popt, self.F)
        self.indices = [i for i in range(self.Z_exp.shape[1])]
        return self.popt, self.perr, self.chisqr, self.chitot, self.AIC

    def fit_stochastic(self,
                       lr: float = 1e-3,
                       num_epochs: int = 1e5,
                       ) -> Tuple[
        torch.tensor, torch.tensor, torch.tensor, torch.tensor, torch.tensor
    ]:  # Optimal parameters, parameter error,
        # weighted residual mean square, and the AIC

        """
        Fitting routine which uses the Adam optimizer.
        It is important to note here that even stocahstic search procedures,
        although applicable to large scale problems do not \
        find the global optimum with certainty (Aster, Richard pg 249)

        :param lr: Learning rate

        :param num_epochs: Number of epochs

        :returns: A tuple containing the optimal parameters (popt), \
                  the standard error of the parameters (perr), \
                  the objective function at the minimum (chisqr), \
                  the total cost function (chitot) and the AIC
        """
        self.lr = lr
        self.num_epochs = int(num_epochs)

        if hasattr(self, "popt") and self.popt.shape[1] == self.Z.shape[1]:
            print("\nUsing prefit")
            self.par_log = (
                self.convert_to_internal(self.check_nan_values(self.popt)).type(torch.DoubleTensor)
            ).requires_grad_(True)
        else:
            print("\nUsing initial")
            self.par_log = (
                self.convert_to_internal(self.p0).type(torch.DoubleTensor)
            ).requires_grad_(True)

        start = datetime.now()
        optimizer = torch.optim.Adam(params=[self.par_log], lr=self.lr)
        self.losses = []
        for epoch in range(self.num_epochs):

            optimizer.zero_grad()
            # loss = get_loss(param,f_data,Z_data)
            self.loss = self.compute_total_obj(
                self.par_log,
                self.F,
                self.Z,
                self.Zerr_Re,
                self.Zerr_Im,
                self.lb_vec,
                self.ub_vec,
                self.smf_1,
            )

            if epoch % int(self.num_epochs / 10) == 0:
                print(
                    ""
                    + str(epoch)
                    + ": "
                    + "loss="
                    + "{:5.3e}".format(self.loss.detach().clone()/self.dof)
                )
            self.losses.append(self.loss.detach().clone()/self.dof)
            self.loss.backward()
            optimizer.step()

        self.popt = self.convert_to_external(self.par_log).detach()
        self.chitot = self.losses[-1]

        self.perr = self.compute_perr(
            self.popt,
            self.F,
            self.Z,
            self.Zerr_Re,
            self.Zerr_Im,
            self.lb_vec,
            self.ub_vec,
            self.smf_1,
        )

        self.chisqr = torch.mean(
            functorch.vmap(self.compute_wrms, in_dims=(1, None, 1, 1, 1))(
                self.popt, self.F, self.Z, self.Zerr_Re, self.Zerr_Im
            )
        )
        self.AIC = torch.mean(
            functorch.vmap(self.compute_aic, in_dims=(1, None, 1, 1, 1))(
                self.popt, self.F, self.Z, self.Zerr_Re, self.Zerr_Im
            )
        )
        print("Optimization complete")
        end = datetime.now()
        print(f"total time is {end-start}", end=" ")
        self.Z_exp = self.Z.clone()
        self.Y_exp = 1 / self.Z_exp.clone()
        self.Z_pred, self.Y_pred = self.model_prediction(self.popt, self.F)
        self.indices = [i for i in range(self.Z_exp.shape[1])]

        return self.popt, self.perr, self.chisqr, self.chitot, self.AIC

    def fit_simultaneous_zero(self,
                              method: str = 'bfgs',
                              n_iter: int = 5000,
                              ) -> Tuple[
        torch.tensor, torch.tensor, torch.tensor, torch.tensor, torch.tensor
    ]:
        """
        Fitting routine with the smoothing factor set to zero.

        :params method: Solver to use (must be one of "'TNC', \
                        'BFGS' or 'L-BFGS-B'")

        :param n_iter: Number of iterations

        :returns: A tuple containing the optimal parameters (popt), \
                  the standard error of the parameters (perr), \
                  the objective function at the minimum (chisqr), \
                  the total cost function (chitot) and the AIC
        """
        self.method = method.lower()
        assert (self.method in ['tnc', 'bfgs', 'l-bfgs-b']), ("method must be one of "
                                                              "'TNC', 'BFGS' or 'L-BFGS-B'")

        if hasattr(self, "popt") and self.popt.shape[1] == self.Z.shape[1]:
            print("\nUsing prefit")
            self.par_log = (
                self.convert_to_internal(self.check_nan_values(self.popt)).type(torch.DoubleTensor)
            ).requires_grad_(True)
        else:
            print("\nUsing initial")
            self.par_log = (
                self.convert_to_internal(self.p0).type(torch.DoubleTensor)
            ).requires_grad_(True)

        start = datetime.now()
        optimizer = ScipyMinimizer(
            params=[self.par_log],
            method=self.method,
            tol=1e-16,
            options={"maxfun" if self.method == "tnc" else "maxiter": n_iter},
        )
        self.iteration = 0

        def closure():
            optimizer.zero_grad()
            loss = self.compute_total_obj(
                self.par_log,
                self.F,
                self.Z,
                self.Zerr_Re,
                self.Zerr_Im,
                self.lb_vec,
                self.ub_vec,
                torch.zeros(self.num_params),
            )
            loss.backward()
            self.iteration += 1
            print(
                "\rIteration : {}, Loss : {:.5e}"
                .format(self.iteration, loss.item()/self.dof),
                end=""
            )
            if self.iteration % 1000 == 0:
                print("")
            return loss

        optimizer.step(closure)
        self.popt = self.convert_to_external(self.par_log).detach()
        self.chitot = torch.as_tensor(optimizer._result.fun)/self.dof

        self.perr = self.compute_perr_QR(
            self.popt,
            self.F,
            self.Z,
            self.Zerr_Re,
            self.Zerr_Im,
        )

        self.chisqr = (
            torch.mean(
                functorch.vmap(self.compute_wrms, in_dims=(1, None, 1, 1, 1))(
                    self.popt, self.F, self.Z, self.Zerr_Re, self.Zerr_Im
                )
            )
        )
        self.AIC = (
            torch.mean(
                functorch.vmap(self.compute_aic, in_dims=(1, None, 1, 1, 1))(
                    self.popt, self.F, self.Z, self.Zerr_Re, self.Zerr_Im
                )
            )
        )
        print("\nOptimization complete")
        end = datetime.now()
        print(f"total time is {end-start}", end=" ")
        self.Z_exp = self.Z.clone()
        self.Y_exp = 1 / self.Z_exp.clone()
        self.Z_pred, self.Y_pred = self.model_prediction(self.popt, self.F)
        self.indices = [i for i in range(self.Z_exp.shape[1])]
        return self.popt, self.perr, self.chisqr, self.chitot, self.AIC

    def fit_sequential(self,
                       indices: Sequence[
                        int
                        ] = None,
                       ) -> Tuple[
        torch.tensor, torch.tensor, torch.tensor, torch.tensor, torch.tensor
    ]:
        """
        Fits each spectra individually using the L-M least squares method

        :params indices: List containing the indices of spectra to plot. \
                         If set to None, all spectra are fitted sequentially

        :returns: A tuple containing the optimal parameters (popt), \
                  the standard error of the parameters (perr), \
                  the objective function at the minimum (chisqr), \
                  the total cost function (chitot) and the AIC
        """

        if indices:
            assert all(
                i < self.num_eis for i in indices
            ), ("One or more values in the indices list "
                "are greater the number of spectra supplied")
            self.indices = indices
            self.n_fits = len(self.indices)
        elif indices is None:
            self.indices = [i for i in range(self.num_eis)]
            self.n_fits = len(self.indices)

        else:
            raise AttributeError("""
            Please choose the index or indices of spectra to fit""")

        if hasattr(self, "popt") and self.popt.shape[1] == self.Z.shape[1]:
            print("\nUsing prefit")
            self.par_log = \
                self.convert_to_internal(self.check_nan_values(self.popt)).type(torch.DoubleTensor)
        else:
            print("\nUsing initial")
            self.par_log = \
                self.convert_to_internal(self.p0).type(torch.DoubleTensor)

        popt = torch.zeros(self.num_params, self.n_fits)
        perr = torch.zeros(self.num_params, self.n_fits)
        params_init = torch.zeros(self.num_params, self.num_eis)
        for i in range(self.num_params):
            params_init[i, :] = (self.par_log)[self.kvals[i]:self.kvals[i + 1]]
        chisqr = torch.zeros(self.n_fits)
        aic = torch.zeros(self.n_fits)
        start = datetime.now()
        for i, val in enumerate(self.indices):
            if i % 10 == 0:
                print(
                    f"fitting spectra {val}"
                )
            try:
                pfit, chi2 = self.do_minimize_ls(
                    params_init[:, val],
                    self.F,
                    self.Z[:, val],
                    self.Zerr_Re[:, val],
                    self.Zerr_Im[:, val],
                    self.lb,
                    self.ub,
                )
            except ValueError:
                pfit = self.encode(params_init[:, val], self.lb, self.ub)
                chi2 = self.compute_rss(
                    params_init[:, val],
                    self.F,
                    self.Z[:, val],
                    self.Zerr_Re[:, val],
                    self.Zerr_Im[:, val],
                    self.lb,
                    self.ub,
                )

            popt[:, i] = self.decode(pfit, self.lb, self.ub).detach()
            chisqr[i] = (
                torch.sum((2*chi2)**2) / (2 * self.num_freq - self.num_params)
                if torch.is_tensor(chi2)
                else torch.sum(torch.tensor(chi2**2))
                / (2 * self.num_freq - self.num_params)
            )

            aic[i] = self.compute_aic(
                popt[:, i],
                self.F,
                self.Z[:, val],
                self.Zerr_Re[:, val],
                self.Zerr_Im[:, val],
            )
            jac = self.compute_jac(
                params_init[:, val],
                self.F,
                self.Z[:, val],
                self.Zerr_Re[:, val],
                self.Zerr_Im[:, val],
                self.lb,
                self.ub,
            )
            hess = jac.T @ jac
            try:
                hess_inv = torch.linalg.inv(hess)
                self.cov_mat_single = hess_inv * (chisqr[i])
                perr[:, i] = torch.sqrt(torch.diag(self.cov_mat_single)) * popt[:, i]
            except torch.linalg.LinAlgError:
                print(
                    "Matrix is singular for spectra {}, using QR decomposition"
                    .format(val)
                )
                grads = torch.autograd.functional.jacobian(
                    self.func, (popt[:, i], self.F)
                )[0]
                gradsre = grads[:self.num_freq]
                gradsim = grads[self.num_freq:]
                diag_wtre_matrix = torch.diag((1 / self.Zerr_Re[:, val]))
                diag_wtim_matrix = torch.diag((1 / self.Zerr_Im[:, val]))
                vre = diag_wtre_matrix.double() @ gradsre.double()
                vim = diag_wtim_matrix.double() @ gradsim.double()
                Q1, R1 = torch.linalg.qr(torch.cat([vre, vim], dim=0))
                try:
                    invR1 = torch.linalg.inv(R1)
                    perr[:, i] = \
                        torch.linalg.vector_norm(invR1, dim=1) * torch.sqrt(
                        chisqr[i]
                    )
                except torch.linalg.LinAlgError:
                    print(
                        """Matrix is singular for spectra {},
                        perr will be assigned a value of ones"""
                        .format(val)
                    )
                    invR1 = torch.linalg.inv(R1)
                    perr[:, i] = \
                        torch.linalg.vector_norm(invR1, dim=1) * torch.sqrt(
                        chisqr[i]
                    )

        self.popt = popt.clone()
        self.perr = torch.nan_to_num(perr.clone(), nan=1.0e15)
        self.chisqr = torch.mean(chisqr)
        self.chitot = self.chisqr.clone()
        self.AIC = torch.mean(aic)
        self.Z_exp = self.Z[:, self.indices]
        self.Y_exp = 1 / self.Z_exp.clone()
        self.Z_pred, self.Y_pred = self.model_prediction(self.popt, self.F)
        print("\nOptimization complete")
        end = datetime.now()
        print(f"total time is {end-start}", end=" ")
        return self.popt, self.perr, self.chisqr, self.chitot, self.AIC

    def compute_perr_mc(self,
                        n_boots: int = 500
                        ) -> Tuple[
        torch.tensor, torch.tensor, torch.tensor, torch.tensor, torch.tensor
    ]:
        """
        The bootstrap approach used here is \
        similar to the fixed-X resampling. \
        In this approach we construct bootstrap observations \
        from the fitted values and the residuals. \
        The assumption that the functional form of the model \
        is implicit in this method. We also assume that \
        the errors are identically distributed with constant variance.
        (Bootstrapping Regression Models - J Fox)

        :param n_boots: Number of bootstrap samples

        :returns: A tuple containing the optimal parameters (popt), \
                  the standard error of the parameters (perr), \
                  the objective function at the minimum (chisqr), \
                  the total cost function (chitot) and the AIC
        """
        print("""\nPlease run fit_simultaneous() or fit_stochastic()
              on your data before running the compute_perr_mc() method.
              ignore this message if you did.""")

        if hasattr(self, "popt") and self.popt.shape[1] == self.Z.shape[1]:
            print("\nUsing prefit")
            par_log = (
                self.convert_to_internal(self.check_nan_values(self.popt)).type(torch.DoubleTensor)
            ).requires_grad_(True)
        else:
            raise ValueError(
                """Please run fit_deterministic(), fit_refine()
                or fit_stochastic() before using
                the compute_perr_mc() method"""
            )

        self.n_boots = n_boots
        wrms = functorch.vmap(self.compute_wrms, in_dims=(1, None, 1, 1, 1))(
            self.popt, self.F, self.Z, self.Zerr_Re, self.Zerr_Im
        )

        self.Z_pred, self.Y_pred = self.model_prediction(self.popt, self.F)

        # Taking the sqrt of the chisquare gives us an
        # estimate of the error in measured immittance values
        rnd_resid_Re = \
            torch.randn(self.num_freq, self.num_eis) * torch.sqrt(wrms)
        rnd_resid_Im = \
            torch.randn(self.num_freq, self.num_eis) * torch.sqrt(wrms)

        if self.weight_name == "sigma":
            Zerr_Re_mc = self.Zerr_Re
            Zerr_Im_mc = self.Zerr_Im
        elif self.weight_name == "proportional":
            Zerr_Re_mc = self.Z_pred.real
            Zerr_Im_mc = self.Z_pred.imag
        elif self.weight_name == "modulus":
            Zerr_Re_mc = torch.abs(self.Z_pred)
            Zerr_Im_mc = torch.abs(self.Z_pred)
        else:
            Zerr_Re_mc = torch.ones(self.num_freq, self.num_eis)
            Zerr_Im_mc = torch.ones(self.num_freq, self.num_eis)

        idx = [i for i in range(self.num_freq)]

        # Make containers to hold bootstrapped values
        self.popt_mc = torch.zeros(self.n_boots, self.num_params, self.num_eis)
        self.Z_pred_mc_tot = torch.zeros(
            size=(self.n_boots, self.num_freq, self.num_eis),
            dtype=torch.complex64
        )
        self.chisqr_mc = torch.zeros(self.n_boots)
        popt_log_mc = torch.zeros(
            self.n_boots,
            self.num_params * self.num_eis
            - (self.num_eis - 1) * torch.sum(torch.isinf(self.smf)),
        )

        # Here we loop through the number of boots and
        # run the minimization algorithm using the do_minimize function
        start = datetime.now()
        par_log_mc = par_log.clone()
        for i in range(self.n_boots):
            sidx = np.random.choice(idx, replace=True, size=self.num_freq)
            rnd_resid_Re_boot = rnd_resid_Re[sidx, :]
            rnd_resid_Im_boot = rnd_resid_Im[sidx, :]
            Z_pred_mc = (
                self.Z_pred.real
                + Zerr_Re_mc * rnd_resid_Re_boot
                + 1j * (self.Z_pred.imag + Zerr_Im_mc * rnd_resid_Im_boot)
            )

            res = self.do_minimize(
                par_log_mc,
                self.F,
                Z_pred_mc,
                Zerr_Re_mc,
                Zerr_Im_mc,
                self.lb_vec,
                self.ub_vec,
                self.smf_1,
            )

            popt_log_mc[i, :] = res.x
            self.popt_mc[i, :, :] = (
                self.convert_to_external(popt_log_mc[i, :])
            ).detach()
            self.chisqr_mc[i] = res.fun/self.dof
            self.Z_pred_mc_tot[i, :, :] = Z_pred_mc
        self.popt = self.popt.clone()
        self.perr = torch.std(self.popt_mc, unbiased=False, dim=0)
        self.chisqr = torch.mean(self.chisqr_mc, dim=0)
        self.chitot = self.chisqr.clone()
        print("\nOptimization complete")
        end = datetime.now()
        print(f"total time is {end-start}", end=" ")
        return self.popt, self.perr, self.chisqr, self.chitot, self.AIC

    def do_minimize(self,
                    p: torch.tensor,
                    f: torch.tensor,
                    z: torch.tensor,
                    zerr_re: torch.tensor,
                    zerr_im: torch.tensor,
                    lb: torch.tensor,
                    ub: torch.tensor,
                    smf: torch.tensor
                    ):
        """

        Fitting routine used in the bootstrap Monte Carlo procedure


        :param p: A 1D tensor of parameter values

        :param f: A 1D tensor of frequency

        :param z: A 1D tensor of complex immittances

        :param zerr_re: A 1D tensor of weights for \
                        the real part of the immittance

        :param zerr_im: A 1D tensor of weights for \
                        the imaginary part of the immittance

        :param lb: A 1D tensor of values for \
                   the lower bounds (for the total parameters)

        :param ub: A 1D tensor of values for \
                   the upper bounds (for the total parameters)

        :param smf: A tensor of real elements same size as p0. \
            when set to inf, the corresponding parameter is kept constant
        """
        res = minimize(
            lambda p0: self.compute_total_obj(p0, f, z, zerr_re, zerr_im, lb, ub, smf),
            p,
            method="bfgs",
            max_iter=5000,
        )
        return res

    def compute_jac(self,
                    p: torch.tensor,
                    f: torch.tensor,
                    z: torch.tensor,
                    zerr_re: torch.tensor,
                    zerr_im: torch.tensor,
                    lb: torch.tensor,
                    ub: torch.tensor,
                    ) -> torch.tensor:
        """
        Computes the Jacobian of the least squares \
        objective function w.r.t the parameters

        :param p: A 1D tensor of parameter values

        :param f: A 1D tensor of frequency

        :param z: A 1D tensor of complex immittance

        :param zerr_re: A 1D tensor of weights for \
                        the real part of the immittance

        :param zerr_im: A 1D tensor of weights for \
                        the imaginary part of the immittance

        :param lb: A 1D tensor of values for \
                   the lower bounds (for the total parameters)

        :param ub: A 1D tensor of values for \
                   the upper bounds (for the total parameters)

        :returns:  Returns the Jacobian matrix
        """
        return torch.autograd.functional.jacobian(
            self.compute_rss, (p, f, z, zerr_re, zerr_im, lb, ub)
        )[0]

    def do_minimize_ls(self,
                       p: torch.tensor,
                       f: torch.tensor,
                       z: torch.tensor,
                       zerr_re: torch.tensor,
                       zerr_im: torch.tensor,
                       lb: torch.tensor,
                       ub: torch.tensor,
                       ) -> Tuple[
        torch.tensor, torch.tensor
    ]:  #
        """
        Least squares routine - uses compute_rss

        :param p: A 1D tensor of parameter values

        :param f: A 1D tensor of frequency

        :param z: A 1D tensor of complex immittance

        :param zerr_re: A 1D tensor of weights for \
                        the real part of the immittance

        :param zerr_im: A 1D tensor of weights for \
                        the imaginary part of the immittance

        :param lb: A 1D tensor of values for \
                   the lower bounds (for the total parameters)

        :param ub: A 1D tensor of values for \
                   the upper bounds (for the total parameters)

        :returns: Returns the log-scaled optimal parameters \
                  and the weighted residual mean square
        """
        res = least_squares(
            lambda p0: self.compute_rss(p0, f, z, zerr_re, zerr_im, lb, ub), p
        )
        return res.x, res.fun

    def encode(self,
               p: torch.tensor,
               lb: torch.tensor,
               ub: torch.tensor,
               ) -> torch.tensor:
        """
        Converts external parameters to internal parameters

        :param p: A 1D tensor of parameter values

        :param lb: A 1D tensor of values for \
                   the lower bounds (for the total parameters)

        :param ub: A 1D tensor of values for \
                   the upper bounds (for the total parameters)

        :returns:  Returns the parameter vector \
                   in log scale (internal coordinates)
        """
        p = torch.log10((p - lb) / (1 - p / ub))
        return p

    def decode(self,
               p: torch.tensor,
               lb: torch.tensor,
               ub: torch.tensor
               ) -> torch.tensor:
        """
        Converts internal parameters to external parameters

        :param p: A 1D tensor of parameter values

        :param lb: A 1D tensor of values for \
                   the lower bounds (for the total parameters)

        :param ub: A 1D tensor of values for \
                   the upper bounds (for the total parameters)

        :returns:  Returns the parameter vector \
                   in normal scale (external coordinates)
        """
        p = (lb + 10 ** (p)) / (1 + 10 ** (p) / ub)
        return p

    def real_to_complex(self,
                        z: torch.tensor,
                        ) -> torch.tensor:
        """
        :param z: real vector of length 2n \
                  where n is the number of frequencies

        :returns: Returns a complex vector of length n.
        """
        return z[: len(z) // 2] + 1j * z[len(z) // 2:]

    def complex_to_real(self,
                        z: torch.tensor,
                        ) -> torch.tensor:

        """
        :param z: complex vector of length n \
                  where n is the number of frequencies

        :returns: Returns a real vector of length 2n
        """

        return torch.cat((z.real, z.imag), dim=0)

    def model_prediction(self,
                         P: torch.tensor,
                         F: torch.tensor
                         ) -> Tuple[torch.tensor, torch.tensor]:
        """
        Computes the predicted immittance and its inverse

        :param P: A 2D tensor of optimal parameters

        :param Z: A 1D tensor of complex immittances

        :returns: The predicted immittance (Z_pred) \
                  and its inverse(Y_pred)
        """
        Z_pred = functorch.vmap(self.real_to_complex, in_dims=0)(
            functorch.vmap(self.func, in_dims=(1, None))(P, F)
        ).T
        Y_pred = 1 / Z_pred.clone()
        return Z_pred, Y_pred

    def create_dir(self,
                   dir_name: str
                   ):
        """
        Creates a directory. equivalent to using mkdir -p on the command line

        :param dir_name: Name assigned to the directory
        """
        self.dir_name = dir_name
        from errno import EEXIST
        from os import makedirs, path

        try:
            makedirs(self.dir_name)
        except OSError as exc:  # Python >2.5
            if exc.errno == EEXIST and path.isdir(self.dir_name):
                pass
            else:
                raise

    def plot_nyquist(self,
                     steps: int = 1,
                     **kwargs,
                     ):
        """
        Creates the complex plane plots (aka Nyquist plots)

        :param steps: Spacing between plots. Defaults to 1

        :keyword fpath1: Additional keyword arguments \
                         passed to plot (i.e file path)

        :keyword fpath2: Additional keyword arguments \
                    passed to plot (i.e file path)

        :returns: The complex plane plots.

        """

        self.steps = steps
        assert (
            self.steps <= self.Z_exp.shape[1]
        ), (
            """Steps with size {} is greater that
            the number of fitted spectra with size {}"""
            .format(steps, self.Z_exp.shape[1]))

        # If the fit method has not been called,
        # only the plots of the experimental data are presented
        if not hasattr(self, "popt"):
            indices = [i for i in range(self.Z_exp.shape[1])]

            self.n_plots = len(
                torch.arange(0, int(self.Z_exp.shape[1]), self.steps)
            )  # (or however many you programatically figure out you need)
            n_cols = 4
            n_rows = 5

            # If self.immittance is impedance then fig_nyquist1 is the
            # impedance plot while fig_nyquist2 is the admittance plot
            self.fig_nyquist1 = plt.figure(figsize=(15, 12), facecolor="white")

            if self.immittance == "impedance":
                # make a plot of the impedance
                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    ax.plot(
                        self.Z_exp[:, i].real,
                        -self.Z_exp[:, i].imag,
                        "o",
                        color="orange",
                        label="Expt",
                    )
                    ax.set(adjustable="datalim", aspect="equal")
                    ax.set_xlabel(r"$\Re(Z)$" + "[" + r"$\Omega$" + "]")
                    ax.set_ylabel(r"$-\Im(Z)$" + "[" + r"$\Omega$" + "]")
                    ax.legend(loc="best")
                    plt.xticks(rotation=45)
                    if k > n_rows * n_cols - 1:
                        break

            else:
                # Make a plot of the admittance
                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    ax.plot(
                        self.Z_exp[:, i].real,
                        self.Z_exp[:, i].imag,
                        "o",
                        color="orange",
                        label="Expt",
                    )
                    ax.set(adjustable="datalim", aspect="equal")
                    ax.set_xlabel(r"$\Re(Y)$" + "[" + r"$S$" + "]")
                    ax.set_ylabel(r"$\Im(Y)$" + "[" + r"$S$" + "]")
                    ax.legend(loc="best")
                    plt.xticks(rotation=45)
                    if k > n_rows * n_cols - 1:
                        break

            self.fig_nyquist1.suptitle(self.plot_title1, y=1.02)
            self.fig_nyquist1.tight_layout(w_pad=0.5, h_pad=0.5)
            if kwargs:
                fpath1 = kwargs.get("fpath1", None)
                self.fig_nyquist1.savefig(
                    fpath1, dpi=300, transparent=False, bbox_inches="tight"
                )
                plt.close(self.fig_nyquist1)
            else:
                plt.show()

            self.fig_nyquist2 = plt.figure(figsize=(15, 12), facecolor="white")

            if self.immittance == "impedance":
                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    ax.plot(
                        self.Y_exp[:, i].real,
                        self.Y_exp[:, i].imag,
                        "o",
                        color="orange",
                        label="Expt",
                    )
                    ax.set(adjustable="datalim", aspect="equal")
                    ax.set_xlabel(r"$\Re(Y)$" + "[" + r"$S$" + "]")
                    ax.set_ylabel(r"$\Im(Y)$" + "[" + r"$S$" + "]")
                    ax.legend(loc="best")
                    plt.xticks(rotation=45)
                    if k > n_rows * n_cols - 1:
                        break

            else:
                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    ax.plot(
                        self.Y_exp[:, i].real,
                        -self.Y_exp[:, i].imag,
                        "o",
                        color="orange",
                        label="Expt",
                    )
                    ax.set(adjustable="datalim", aspect="equal")
                    ax.set_xlabel(r"$\Re(Z)$" + "[" + r"$\Omega$" + "]")
                    ax.set_ylabel(r"$-\Im(Z)$" + "[" + r"$\Omega$" + "]")
                    ax.legend(loc="best")
                    plt.xticks(rotation=45)
                    if k > n_rows * n_cols - 1:
                        break

            self.fig_nyquist2.suptitle(self.plot_title2, y=1.02)
            self.fig_nyquist2.tight_layout(w_pad=0.5, h_pad=0.5)
            if kwargs:
                fpath2 = kwargs.get("fpath2", None)
                self.fig_nyquist2.savefig(
                    fpath2, dpi=300, transparent=False, bbox_inches="tight"
                )
                plt.close(self.fig_nyquist2)
            else:
                plt.show()

        else:
            # If a fit method has been called then assign self.indices to indices
            indices = self.indices

            n_cols = 4
            n_rows = 5

            self.fig_nyquist1 = plt.figure(figsize=(15, 12), facecolor="white")

            if self.immittance == "impedance":

                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    ax.plot(
                        self.Z_exp[:, i].real,
                        -self.Z_exp[:, i].imag,
                        "o",
                        color="orange",
                        label="Expt",
                    )
                    ax.plot(
                        (self.Z_pred[:, i]).real,
                        -(self.Z_pred[:, i]).imag,
                        "b-",
                        lw=1.5,
                        label="Model",
                    )
                    ax.set(adjustable="datalim", aspect="equal")
                    ax.set_xlabel(r"$\Re(Z)$" + "[" + r"$\Omega$" + "]")
                    ax.set_ylabel(r"$-\Im(Z)$" + "[" + r"$\Omega$" + "]")
                    ax.legend(loc="best")
                    plt.xticks(rotation=45)
                    if k > n_rows * n_cols - 1:
                        break

            else:
                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    ax.plot(
                        self.Z_exp[:, i].real,
                        self.Z_exp[:, i].imag,
                        "o",
                        color="orange",
                        label="Expt",
                    )
                    ax.plot(
                        (self.Z_pred[:, i]).real,
                        (self.Z_pred[:, i]).imag,
                        "b-",
                        lw=1.5,
                        label="Model",
                    )
                    ax.set(adjustable="datalim", aspect="equal")
                    ax.set_xlabel(r"$\Re(Y)$" + "[" + r"$S$" + "]")
                    ax.set_ylabel(r"$\Im(Y)$" + "[" + r"$S$" + "]")
                    ax.legend(loc="best")
                    plt.xticks(rotation=45)
                    if k > n_rows * n_cols - 1:
                        break

            self.fig_nyquist1.suptitle(self.plot_title1, y=1.02)
            self.fig_nyquist1.tight_layout(w_pad=0.5, h_pad=0.5)
            if kwargs:
                fpath1 = kwargs.get("fpath1", None)
                self.fig_nyquist1.savefig(
                    fpath1, dpi=300, transparent=False, bbox_inches="tight"
                )
                plt.close(self.fig_nyquist1)
            else:
                plt.show()

            self.fig_nyquist2 = plt.figure(figsize=(15, 12), facecolor="white")

            if self.immittance == "impedance":
                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    ax.plot(
                        self.Y_exp[:, i].real,
                        self.Y_exp[:, i].imag,
                        "o",
                        color="orange",
                        label="Expt",
                    )
                    ax.plot(
                        (self.Y_pred[:, i]).real,
                        (self.Y_pred[:, i]).imag,
                        "b-",
                        lw=1.5,
                        label="Model",
                    )
                    ax.set(adjustable="datalim", aspect="equal")
                    ax.set_xlabel(r"$\Re(Y)$" + "[" + r"$S$" + "]")
                    ax.set_ylabel(r"$\Im(Y)$" + "[" + r"$S$" + "]")
                    ax.legend(loc="best")
                    plt.xticks(rotation=45)
                    if k > n_rows * n_cols - 1:
                        break

            else:
                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    ax.plot(
                        self.Y_exp[:, i].real,
                        -self.Y_exp[:, i].imag,
                        "o",
                        color="orange",
                        label="Expt",
                    )
                    ax.plot(
                        (self.Y_pred[:, i]).real,
                        -(self.Y_pred[:, i]).imag,
                        "b-",
                        lw=1.5,
                        label="Model",
                    )
                    ax.set(adjustable="datalim", aspect="equal")
                    ax.set_xlabel(r"$\Re(Z)$" + "[" + r"$\Omega$" + "]")
                    ax.set_ylabel(r"$-\Im(Z)$" + "[" + r"$\Omega$" + "]")
                    ax.legend(loc="best")
                    plt.xticks(rotation=45)
                    if k > n_rows * n_cols - 1:
                        break

            self.fig_nyquist2.suptitle(self.plot_title2, y=1.02)
            self.fig_nyquist2.tight_layout(w_pad=0.5, h_pad=0.5)
            if kwargs:
                fpath2 = kwargs.get("fpath2", None)
                self.fig_nyquist2.savefig(
                    fpath2, dpi=300, transparent=False, bbox_inches="tight"
                )
                plt.close(self.fig_nyquist2)
            else:
                plt.show()

    def plot_bode(self,
                  steps: int = 1,
                  **kwargs,
                  ):
        """
        Creates the Bode plots
        The Bode plot shows the phase angle of a
        capacitor's or inductors opptosition to current.
        A capacitor's opposition to current is -90°,
        which means that a capacitor's opposition
        to current is a negative imaginary quantity.


        :param steps: Spacing between plots. Defaults to 1.

        :keyword fpath: Additional keyword arguments \
                         passed to plot (i.e file path)

        :returns: The bode plots.

        Notes
        ---------

        .. math::

            \\theta = arctan2 (\\frac{\\Im{Z}}{\\Re{Z}} \\frac{180}{\\pi})

        """
        assert (
            steps <= self.Z_exp.shape[1]
        ), (
            """Steps with size {} is greater that the
            number of fitted spectra with size {}"""
            .format(steps, self.Z_exp.shape[1]))
        self.steps = steps
        self.Z_mag = torch.abs(self.Z_exp)
        self.Z_angle = torch.rad2deg(torch.atan2(self.Z_exp.imag, self.Z_exp.real))

        if not hasattr(
            self, "popt"
        ):  # If the fit method has not been called,
            # only the plots of the experimental data are presented

            indices = [
                i for i in range(self.Z_exp.shape[1])
            ]
            # Indices should be determined by Z_exp
            # which changes depending on the routine used
            n_cols = 4
            n_rows = 5

            self.fig_bode = plt.figure(figsize=(15, 12), facecolor="white")
            if self.immittance == "impedance":

                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax1 = ax.twinx()
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    (p1,) = ax.plot(
                        self.F,
                        self.Z_mag[:, i],
                        "o",
                        markersize=5,
                        color="blue",
                        label="Expt",
                    )
                    (p2,) = ax1.plot(
                        self.F,
                        self.Z_angle[:, i],
                        "o",
                        markersize=5,
                        color="orange",
                        label="Expt",
                    )
                    ax.set(adjustable="datalim")
                    ax.set_xlim(self.F.min(), self.F.max())
                    ax1.invert_yaxis()
                    ax.relim()
                    ax.autoscale()
                    ax.semilogx()
                    ax.semilogy()
                    ax.set_xlabel("$F$" + "[" + "$Hz$" + "]")
                    ax.set_ylabel(r"$|Z|$" + "[" + r"$\Omega$" + "]")
                    ax1.set_ylabel(r"$\theta$ $(^{\circ}) $")
                    ax.yaxis.label.set_color(p1.get_color())
                    ax1.yaxis.label.set_color(p2.get_color())
                    ax.tick_params(axis="y", colors=p1.get_color(), which="both")
                    ax1.tick_params(axis="y", colors=p2.get_color(), which="both")
                    if k > n_rows * n_cols - 1:
                        break
            else:
                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax1 = ax.twinx()
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    (p1,) = ax.plot(
                        self.F,
                        self.Z_mag[:, i],
                        "o",
                        markersize=5,
                        color="blue",
                        label="Expt",
                    )
                    (p2,) = ax1.plot(
                        self.F,
                        -self.Z_angle[:, i],
                        "o",
                        markersize=5,
                        color="orange",
                        label="Expt",
                    )
                    ax.set(adjustable="datalim")
                    ax.set_xlim(self.F.min(), self.F.max())
                    ax1.invert_yaxis()
                    ax.relim()
                    ax.autoscale()
                    ax.semilogx()
                    ax.semilogy()
                    ax.set_xlabel("$F$" + "[" + "$Hz$" + "]")
                    ax.set_ylabel("$|Y|$" + "[" + "$S$" + "]")
                    ax1.set_ylabel(r"$\theta$ $(^{\circ}) $")
                    ax.yaxis.label.set_color(p1.get_color())
                    ax1.yaxis.label.set_color(p2.get_color())
                    ax.tick_params(axis="y", colors=p1.get_color(), which="both")
                    ax1.tick_params(axis="y", colors=p2.get_color(), which="both")
                    if k > n_rows * n_cols - 1:
                        break

            self.fig_bode.suptitle("Bode Plot", y=1.02)
            plt.tight_layout(w_pad=0.5, h_pad=0.5)
            if kwargs:
                fpath = kwargs.get("fpath", None)
                self.fig_bode.savefig(
                    fpath, dpi=300, transparent=False, bbox_inches="tight"
                )
                plt.close(self.fig_bode)
            else:
                plt.show()

        else:
            indices = (
                self.indices
            )  # Indices becomes sel.indices after a fit method has been called.
            self.Z_mag_pred = torch.abs(self.Z_pred)
            self.Z_angle_pred = torch.rad2deg(
                torch.atan2(self.Z_pred.imag, self.Z_pred.real)
            )

            self.n_plots = len(
                torch.arange(0, int(self.Z_exp.shape[1]), self.steps)
            )  # (or however many you programatically figure out you need)
            n_cols = 4
            n_rows = 5

            self.fig_bode = plt.figure(figsize=(15, 12), facecolor="white")
            if self.immittance == "impedance":

                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax1 = ax.twinx()
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    (p1,) = ax.plot(
                        self.F,
                        self.Z_mag[:, i],
                        "o",
                        markersize=5,
                        color="blue",
                        label="Expt",
                    )
                    ax.plot(
                        self.F,
                        self.Z_mag_pred[:, i],
                        "-",
                        color="red",
                        lw=1.5,
                        label="Model",
                    )
                    (p2,) = ax1.plot(
                        self.F,
                        self.Z_angle[:, i],
                        "o",
                        markersize=5,
                        color="orange",
                        label="Expt",
                    )
                    ax1.plot(
                        self.F,
                        self.Z_angle_pred[:, i],
                        "-",
                        color="purple",
                        lw=1.5,
                        label="Model",
                    )
                    ax.set(adjustable="datalim")
                    ax.set_xlim(self.F.min(), self.F.max())
                    ax1.invert_yaxis()
                    ax.relim()
                    ax.autoscale()
                    ax.semilogx()
                    ax.semilogy()
                    ax.set_xlabel("$F$" + "[" + "$Hz$" + "]")
                    ax.set_ylabel(r"$|Z|$" + "[" + r"$\Omega$" + "]")
                    ax1.set_ylabel(r"$\theta$ $(^{\circ}) $")
                    ax.yaxis.label.set_color(p1.get_color())
                    ax1.yaxis.label.set_color(p2.get_color())
                    ax.tick_params(axis="y", colors=p1.get_color(), which="both")
                    ax1.tick_params(axis="y", colors=p2.get_color(), which="both")
                    if k > n_rows * n_cols - 1:
                        break
            else:
                k = 0
                for i in range(0, self.Z_exp.shape[1], self.steps):

                    k = k + 1
                    ax = plt.subplot(n_rows, n_cols, k)
                    ax1 = ax.twinx()
                    ax.set_title("Idx = {:d}".format(indices[i]))
                    (p1,) = ax.plot(
                        self.F,
                        self.Z_mag[:, i],
                        "o",
                        markersize=5,
                        color="blue",
                        label="Expt",
                    )
                    ax.plot(
                        self.F,
                        self.Z_mag_pred[:, i],
                        "-",
                        color="red",
                        lw=1.5,
                        label="Model",
                    )
                    (p2,) = ax1.plot(
                        self.F,
                        -self.Z_angle[:, i],
                        "o",
                        markersize=5,
                        color="orange",
                        label="Expt",
                    )
                    ax1.plot(
                        self.F,
                        -self.Z_angle_pred[:, i],
                        "-",
                        color="purple",
                        lw=1.5,
                        label="Model",
                    )
                    ax.set(adjustable="datalim")
                    ax.set_xlim(self.F.min(), self.F.max())
                    ax1.invert_yaxis()
                    ax.relim()
                    ax.autoscale()
                    ax.semilogx()
                    ax.semilogy()
                    ax.set_xlabel("$F$" + "[" + "$Hz$" + "]")
                    ax.set_ylabel("$|Y|$" + "[" + "$S$" + "]")
                    ax1.set_ylabel(r"$\theta$ $(^{\circ}) $")
                    ax.yaxis.label.set_color(p1.get_color())
                    ax1.yaxis.label.set_color(p2.get_color())
                    ax.tick_params(axis="y", colors=p1.get_color(), which="both")
                    ax1.tick_params(axis="y", colors=p2.get_color(), which="both")
                    if k > n_rows * n_cols - 1:
                        break

            self.fig_bode.suptitle("Bode Plot", y=1.02)
            plt.tight_layout(w_pad=0.5, h_pad=0.5)
            if kwargs:
                fpath = kwargs.get("fpath", None)
                self.fig_bode.savefig(
                    fpath, dpi=300, transparent=False, bbox_inches="tight"
                )
                plt.close(self.fig_bode)
            else:
                plt.show()

    def plot_params(self,
                    show_errorbar: bool = False,
                    labels: Dict[str, str] = None,
                    **kwargs,
                    ) -> None:
        """
        Creates the plot of the optimal parameters as a function of the index

        :param show_errorbar: If set to True, \
                              the errorbars are shown on the parameter plot.

        :param labels: A dictionary containing the circuit elements\
                       as keys and the units as values e.g \
                       labels = {
                        "Rs":"$\\Omega$",
                        "Qh":"$F$",
                        "nh":"-",
                        "Rct":"$\\Omega$",
                        "Wct":"$\\Omega\\cdot^{0.5}$",
                        "Rw":"$\\Omega$"
                        }

        :keyword fpath: Additional keyword arguments \
                         passed to plot (i.e file path)

        :returns: The parameter plots
        """
        if labels is None:
            self.labels = [str(i) for i in range(self.num_params)]
        else:

            assert (isinstance(labels, collections.Mapping)), (
                """labels is not a valid dictionary"""
                )

            assert (len(labels.items()) == self.num_params), (
                """Ths size of the labels is {}
                while the size of the parameters is {}"""
                .format(
                    len(labels.items()), self.num_params
                    )
                )

            self.labels = {
                self.try_convert(k): self.try_convert(v) for k, v in labels.items()
                }

        self.show_errorbar = show_errorbar

        if not hasattr(self, "popt"):
            raise AttributeError("A fit() method has not been called.")
        else:

            self.param_idx = [int(i) for i in self.indices]
            params_df = pd.DataFrame(
                self.popt.T.numpy(),
                columns=[i for i in range(self.num_params)]
                )
            params_df['Idx'] = self.param_idx
            params_df['Idx'] = params_df["Idx"].astype('category')
            self.params_df = params_df.fillna(0)
            if self.show_errorbar is True:
                # Plot with error bars
                self.fig_params = (
                    self.params_df.plot(
                        x='Idx',
                        marker="o",
                        linestyle="--",
                        subplots=True,
                        layout=(5, 5),
                        yerr=self.perr.numpy(),
                        figsize=(15, 12),
                        rot=45,
                        legend=False,
                    )
                    .ravel()[0]
                    .get_figure()
                )
            else:
                # Plot without error bars
                self.fig_params = (
                    self.params_df.plot(
                        x='Idx',
                        marker="o",
                        linestyle="--",
                        subplots=True,
                        layout=(5, 5),
                        figsize=(15, 12),
                        rot=45,
                        legend=False
                    )
                    .ravel()[0]
                    .get_figure()
                )
            plt.suptitle("Evolution of parameters ", y=1.01)
            plt.gcf().set_facecolor("white")
            all_axes = plt.gcf().get_axes()
            if labels is not None:
                for i, (k, v) in enumerate(self.labels.items()):
                    all_axes[i].set_ylabel(k + " " + "/" + " " + v, rotation=90)
            else:
                for i, v in enumerate(self.labels):
                    all_axes[i].set_ylabel(v, rotation=90)

            plt.tight_layout()
            if kwargs:
                fpath = kwargs.get("fpath", None)
                self.fig_params.savefig(
                    fpath, dpi=300, transparent=False, bbox_inches="tight"
                )
                plt.close(self.fig_params)
            else:
                plt.show()

    def get_img_path(self,
                     fname:
                     str = None
                     ):
        """
        Creates a path name for saving images
        """
        if fname is None:
            img_path = os.path.join(os.path.abspath(os.getcwd()), "fit")
            img_folder = os.path.join(img_path, "images")
            self.create_dir(img_folder)
            path_name = os.path.join(img_folder, "fit")

        elif isinstance(fname, str):
            img_path = os.path.join(os.path.abspath(os.getcwd()), fname)
            img_folder = os.path.join(img_path, "images")
            self.create_dir(img_folder)
            path_name = os.path.join(img_folder, fname)

        else:
            raise TypeError(
                f"Oops! {fname} is not valid. fname should be None or a valid string"
            )

        return path_name

    def get_results_path(self,
                         fname: str = None
                         ):
        """
        Creates a path name for saving the results
        """
        if fname is None:
            results_path = os.path.join(os.path.abspath(os.getcwd()), "fit")
            results_folder = os.path.join(results_path, "results")
            self.create_dir(results_folder)
            path_name = os.path.join(results_folder, "fit")

        elif isinstance(fname, str):
            results_path = os.path.join(os.path.abspath(os.getcwd()), fname)
            results_folder = os.path.join(results_path, "results")
            self.create_dir(results_folder)
            path_name = os.path.join(results_folder, fname)

        else:
            raise TypeError(
                f"Oops! {fname} is not valid. fname should be None or a valid string"
            )

        return path_name

    def save_plot_nyquist(self,
                          steps: int = 1,
                          *,
                          fname: str = None,
                          ) -> None:
        """
        Saves the Nyquist plots in the current working directory
        with the fname provided.

        :param steps: Spacing between plots. Defaults to 1.

        :keyword fname: Name assigned to the directory, generated plots and data

        :returns: A .png image of the the complex plane plots
        """

        self.img_path_name = self.get_img_path(fname)
        try:
            self.plot_nyquist(
                steps,
                fpath1=self.img_path_name + "_" + self.plot_title1.lower() + ".png",
                fpath2=self.img_path_name + "_" + self.plot_title2.lower() + ".png",
            )

        except AttributeError as e:
            logging.exception("", e, exc_info=True)

    def save_plot_bode(self,
                       steps: int = 1,
                       *,
                       fname: str = None,
                       ) -> None:
        """
        Saves the Bode plots in the current working directory
        with the fname provided

        :param steps: Spacing between plots. Defaults to 1.

        :keyword fname: Name assigned to the directory, generated plots and data

        :returns: A .png image of the the bode plot
        """
        self.img_path_name = self.get_img_path(fname)
        try:
            self.plot_bode(steps, fpath=self.img_path_name + "_bode" + ".png")

        except AttributeError as e:
            logging.exception("", e, exc_info=True)

    def save_plot_params(self,
                         show_errorbar: bool = False,
                         labels: Dict[str, str] = None,
                         *,
                         fname: str = None,
                         ) -> None:
        """
        Saves the parameter plots in the current working directory
        with the fname provided.

        :param show_errorbar: If set to True, \
                              the errorbars are shown on the parameter plot.


        :param labels: A dictionary containing the circuit elements\
                       as keys and the units as values e.g \
                       labels = {
                        "Rs":"$\\Omega$",
                        "Qh":"$F$",
                        "nh":"-",
                        "Rct":"$\\Omega$",
                        "Wct":"$\\Omega\\cdot^{0.5}$",
                        "Rw":"$\\Omega$"
                        }

        :keyword fname: Name assigned to the directory, generated plots and data

        :returns: A .png image of the the parameter plot
        """
        if labels is None:
            self.labels = None
        else:
            assert (isinstance(labels, collections.Mapping)), (
                """labels is not a valid dictionary"""
                )

            assert (len(labels.items()) == self.num_params), (
                """Ths size of the labels is {}
                while the size of the parameters is {}"""
                .format(
                    len(labels.items()), self.num_params
                    )
                )

            self.labels = {
                self.try_convert(k): self.try_convert(v) for k, v in labels.items()
                }

        if not hasattr(self, "popt"):
            raise AttributeError("A fit() method has not been called.")
        else:
            self.img_path_name = self.get_img_path(fname)
            try:
                self.plot_params(
                    show_errorbar=show_errorbar,
                    labels=self.labels,
                    fpath=self.img_path_name + "_params" + ".png"
                )

            except AttributeError as e:
                logging.exception("", e, exc_info=True)

    def save_results(self,
                     *,
                     fname: str = None,
                     ):  # The complex plane, bode and the parameter plots.
        """
        Saves the results (popt, perr, and Z_pred) in the current working directory
        with the fname provided

        :keyword fname: Name assigned to the directory, generated plots and data

        :returns: A .png image of the the complex plane plots
        """
        if not hasattr(self, "popt"):
            raise AttributeError("A fit() method has not been called.")
        else:
            self.results_path_name = self.get_results_path(fname)
            np.save(self.results_path_name + "_popt.npy", self.popt.numpy())
            np.save(self.results_path_name + "_perr.npy", self.perr.numpy())
            np.save(self.results_path_name + "_Z_pred.npy", self.Z_pred.numpy())
            with open(self.results_path_name + "_metrics.txt", "w") as fh:
                fh.write(
                    "%s %s %s %s %s\r"
                    % ("Immittance", "Weight", "AIC", "chisqr", "chitot")
                )
                fh.write(
                    "%s %s %.2f %.2e %.2e\r"
                    % (
                        self.immittance,
                        self.weight_name,
                        self.AIC.numpy(),
                        self.chisqr.numpy(),
                        self.chitot.numpy(),
                    )
                )
