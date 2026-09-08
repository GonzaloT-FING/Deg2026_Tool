function report = rename_first_hash_index(folder_path, from_index, to_index, dry_run, recursive)
%RENAME_FIRST_HASH_INDEX Change only the FIRST "_#number" index in .DTA filenames.
%
% Usage:
%   rename_first_hash_index(folder_path, from_index, to_index)
%   rename_first_hash_index(folder_path, from_index, to_index, dry_run)
%   rename_first_hash_index(folder_path, from_index, to_index, dry_run, recursive)
%
% Defaults:
%   dry_run   = true
%   recursive = false
%
% Example for 1 -> 2:
%   ..._#1.DTA     -> ..._#2.DTA
%   ..._#1_#2.DTA  -> ..._#2_#2.DTA
%   ..._#1_#1.DTA  -> ..._#2_#1.DTA
%   ..._#3_#1.DTA  -> unchanged
%
% Only the FIRST "_#number" token is considered.
% File contents are not modified.

  if nargin < 3
    error(['Usage: rename_first_hash_index(folder_path, from_index, ' ...
           'to_index [, dry_run [, recursive]])']);
  endif

  if isempty(folder_path)
    folder_path = pwd;
  endif

  if nargin < 4 || isempty(dry_run)
    dry_run = true;
  endif

  if nargin < 5 || isempty(recursive)
    recursive = false;
  endif

  if exist(folder_path, 'dir') != 7
    error('Folder does not exist: %s', folder_path);
  endif

  validate_index(from_index, 'from_index');
  validate_index(to_index, 'to_index');

  from_index = round(from_index);
  to_index = round(to_index);

  if from_index == to_index
    error('from_index and to_index are the same (%d).', from_index);
  endif

  files = collect_dta_files(folder_path, recursive);

  old_paths = {};
  new_paths = {};
  old_names = {};
  new_names = {};

  for k = 1:numel(files)
    old_name = files(k).name;

    [starts, ends, tokens] = regexp(old_name, '_#([0-9]+)', ...
                                    'start', 'end', 'tokens');

    if isempty(starts)
      continue;
    endif

    % Only inspect the first _#number token.
    first_index = str2double(tokens{1}{1});

    if first_index != from_index
      continue;
    endif

    new_token = sprintf('_#%d', to_index);

    new_name = [old_name(1:starts(1)-1), ...
                new_token, ...
                old_name(ends(1)+1:end)];

    old_paths{end + 1} = fullfile(files(k).folder, old_name);
    new_paths{end + 1} = fullfile(files(k).folder, new_name);
    old_names{end + 1} = old_name;
    new_names{end + 1} = new_name;
  endfor

  total_matches = numel(old_paths);
  renamed = 0;
  skipped = 0;
  failed = 0;

  fprintf('\nFolder: %s\n', folder_path);
  fprintf('Changing first index: #%d -> #%d\n', from_index, to_index);
  fprintf('Recursive: %s\n', logical_text(recursive));
  fprintf('Mode: %s\n\n', ternary(dry_run, 'PREVIEW ONLY', 'RENAME FILES'));

  if total_matches == 0
    fprintf('No matching .DTA filenames were found.\n');
    report = make_report(total_matches, renamed, skipped, failed, dry_run, ...
                         from_index, to_index);
    return;
  endif

  duplicate_target = false(1, total_matches);

  for i = 1:total_matches
    for j = i + 1:total_matches
      if strcmpi(new_paths{i}, new_paths{j})
        duplicate_target(i) = true;
        duplicate_target(j) = true;
      endif
    endfor
  endfor

  for k = 1:total_matches
    fprintf('%4d. %s\n      -> %s\n', k, old_names{k}, new_names{k});

    if duplicate_target(k)
      fprintf('      SKIPPED: duplicate destination generated.\n');
      skipped = skipped + 1;
      continue;
    endif

    if exist(new_paths{k}, 'file') == 2
      fprintf('      SKIPPED: destination already exists.\n');
      skipped = skipped + 1;
      continue;
    endif

    if dry_run
      continue;
    endif

    [ok, message] = movefile(old_paths{k}, new_paths{k});

    if ok
      renamed = renamed + 1;
    else
      fprintf('      FAILED: %s\n', message);
      failed = failed + 1;
    endif
  endfor

  fprintf('\nSummary\n');
  fprintf('  Matching filenames: %d\n', total_matches);

  if dry_run
    fprintf('  Files renamed:      0 (preview mode)\n');
    fprintf('  Files skipped:      %d\n', skipped);
    fprintf('\nRun again with dry_run = false to apply the changes.\n');
  else
    fprintf('  Files renamed:      %d\n', renamed);
    fprintf('  Files skipped:      %d\n', skipped);
    fprintf('  Failures:           %d\n', failed);
  endif

  report = make_report(total_matches, renamed, skipped, failed, dry_run, ...
                       from_index, to_index);
endfunction


function validate_index(value, name)
  if !isnumeric(value) || !isscalar(value) || !isfinite(value)
    error('%s must be one finite numeric scalar.', name);
  endif

  if value < 0 || value != fix(value)
    error('%s must be a non-negative integer.', name);
  endif
endfunction


function files = collect_dta_files(folder_path, recursive)
  entries = dir(folder_path);
  files = struct('name', {}, 'folder', {});

  for k = 1:numel(entries)
    name = entries(k).name;

    if entries(k).isdir
      if recursive && !strcmp(name, '.') && !strcmp(name, '..')
        subfolder = fullfile(folder_path, name);
        subfiles = collect_dta_files(subfolder, true);
        files = [files, subfiles];
      endif
    elseif !isempty(regexpi(name, '\.DTA$', 'once'))
      item.name = name;
      item.folder = folder_path;
      files(end + 1) = item;
    endif
  endfor
endfunction


function text = logical_text(value)
  if value
    text = 'yes';
  else
    text = 'no';
  endif
endfunction


function result = ternary(condition, true_value, false_value)
  if condition
    result = true_value;
  else
    result = false_value;
  endif
endfunction


function report = make_report(matches, renamed, skipped, failed, dry_run, ...
                              from_index, to_index)
  report = struct( ...
    'matching_filenames', matches, ...
    'renamed', renamed, ...
    'skipped', skipped, ...
    'failed', failed, ...
    'dry_run', logical(dry_run), ...
    'from_index', from_index, ...
    'to_index', to_index);
endfunction
