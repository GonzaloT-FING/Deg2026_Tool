from __future__ import annotations


DEFAULT_LANGUAGE = "es"

LANGUAGE_LABELS = {
    "es": "Español",
    "en": "English",
}

LANGUAGE_CODE_BY_LABEL = {label: code for code, label in LANGUAGE_LABELS.items()}

TRANSLATIONS = {
    "es": {
        "language": "Idioma",
        "language_status": "Idioma activo: {language}.",
        "voltage": "Voltaje",
        "current": "Corriente",
        "temperature": "Temperatura",
        "time": "Tiempo",
        "current_density": "Densidad de corriente",
        "series_by_time": "Series por tiempo",
        "v_vs_i": "V vs I",
        "dv_di": "dV/dI",
        "step_stability": "Estabilidad por paso",
        "ascending": "Ascendente",
        "descending": "Descendente",
        "voltage_short": "V",
        "ascending_finish_marker": "Fin Asc / inicio Dsc",
        "ascending_finish_time": "Tiempo fin ascendente",
        "total_time": "Tiempo total",
        "voltage_range": "Rango de voltaje",
        "average_temperature": "Temperatura promedio",
        "maximum_delta_t": "Delta T maximo",
        "high_current_target": "Corriente objetivo alta",
        "high_current_dvdi": "dV/dI a corriente alta ({direction})",
        "maximum_dvdi": "dV/dI maximo ({direction})",
        "maximum_dvdi_current": "Corriente en dV/dI maximo ({direction})",
        "delta_voltage": "Delta V",
        "delta_temperature": "Delta T",
        "voltage_step_range": "Rango de voltaje por paso",
        "temperature_step_range": "Rango de temperatura por paso",
        "curve": "Curva",
        "technique": "Tecnica",
        "date": "Fecha",
        "start_time": "Hora",
        "current_range": "Rango de corriente",
        "step_duration": "Duracion del paso",
        "current_step": "Paso de corriente",
        "sweeping_rate": "Velocidad de barrido",
        "sampling_time": "Tiempo de muestreo",
        "area": "Area",
        "asc_files": "Archivos Asc",
        "dsc_files": "Archivos Dsc",
        "name": "Nombre",
        "value": "Valor",
        "unit": "Unidad",
        "metadata": "Metadata",
        "indicators": "Indicadores",
        "pc_series_report_title": "Reporte PC Series por tiempo - {curve}",
        "pc_series_report_subtitle": "Metadata e indicadores de la curva de polarizacion",
        "pc_report_title": "Reporte PC - {curve}",
        "pc_report_subtitle": "Metadata e indicadores de la curva de polarizacion",
        "pc_dvdi_report_title": "Reporte PC dV/dI - {curve}",
        "pc_dvdi_report_subtitle": "Metadata e indicadores dV/dI de la curva de polarizacion",
    },
    "en": {
        "language": "Language",
        "language_status": "Active language: {language}.",
        "voltage": "Voltage",
        "current": "Current",
        "temperature": "Temperature",
        "time": "Time",
        "current_density": "Current density",
        "series_by_time": "Series by time",
        "v_vs_i": "V vs I",
        "dv_di": "dV/dI",
        "step_stability": "Step stability",
        "ascending": "Ascending",
        "descending": "Descending",
        "voltage_short": "V",
        "ascending_finish_marker": "End Asc / start Dsc",
        "ascending_finish_time": "Ascending finish time",
        "total_time": "Total time",
        "voltage_range": "Voltage range",
        "average_temperature": "Average temperature",
        "maximum_delta_t": "Maximum Delta T",
        "high_current_target": "High current target",
        "high_current_dvdi": "High current dV/dI ({direction})",
        "maximum_dvdi": "Maximum dV/dI ({direction})",
        "maximum_dvdi_current": "Current at maximum dV/dI ({direction})",
        "delta_voltage": "Delta V",
        "delta_temperature": "Delta T",
        "voltage_step_range": "Voltage range per step",
        "temperature_step_range": "Temperature range per step",
        "curve": "Curve",
        "technique": "Technique",
        "date": "Date",
        "start_time": "Time",
        "current_range": "Current range",
        "step_duration": "Step duration",
        "current_step": "Current step",
        "sweeping_rate": "Sweeping rate",
        "sampling_time": "Sampling time",
        "area": "Area",
        "asc_files": "Asc files",
        "dsc_files": "Dsc files",
        "name": "Name",
        "value": "Value",
        "unit": "Unit",
        "metadata": "Metadata",
        "indicators": "Indicators",
        "pc_series_report_title": "PC Series by time Report - {curve}",
        "pc_series_report_subtitle": "Polarization curve metadata and indicators",
        "pc_report_title": "PC Report - {curve}",
        "pc_report_subtitle": "Polarization curve metadata and indicators",
        "pc_dvdi_report_title": "PC dV/dI Report - {curve}",
        "pc_dvdi_report_subtitle": "Polarization curve dV/dI metadata and indicators",
    },
}


def normalize_language(language: str | None) -> str:
    if language in LANGUAGE_LABELS:
        return str(language)
    return DEFAULT_LANGUAGE


def translate(key: str, lang: str | None = None, **kwargs: object) -> str:
    language_code = normalize_language(lang)
    template = TRANSLATIONS.get(language_code, {}).get(
        key,
        TRANSLATIONS[DEFAULT_LANGUAGE].get(key, key),
    )
    return template.format(**kwargs) if kwargs else template
