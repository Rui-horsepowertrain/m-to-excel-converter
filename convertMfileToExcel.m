function convertMfileToExcel(inputFile, outputFile, decimals)
    % convertMfileToExcel Convert MATLAB .m file constants, tables, and maps to Excel
    %
    % Usage:
    %   convertMfileToExcel('input.m', 'output.xlsx', 3);
    %
    % Inputs:
    %   inputFile  - Path to .m file
    %   outputFile - Path to output .xlsx file
    %   decimals   - Number of decimal places (default 3)
    
    if nargin < 3
        decimals = 3;
    end

    txt = fileread(inputFile);
    txt = strrep(txt, sprintf('\r\n'), sprintf('\n'));
    txt = strrep(txt, sprintf('\r'), sprintf('\n'));

    % Remove comments
    lines = regexp(txt, '\n', 'split');
    cleanTxt = '';
    for i = 1:numel(lines)
        line = lines{i};
        idx = find(line == '%', 1);
        if ~isempty(idx)
            line = line(1:idx-1);
        end
        cleanTxt = [cleanTxt line newline]; %#ok<AGROW>
    end

    constNames = {};
    constValues = [];
    tableNames = {};
    tableValues = {};
    mapNames = {};
    mapValues = {};

    % Match assignments: name = value;
    expr = '([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(\[[\s\S]*?\]|[^;\n]*);';
    toks = regexp(cleanTxt, expr, 'tokens');

    for k = 1:numel(toks)
        name = toks{k}{1};
        value = strtrim(toks{k}{2});

        if isempty(value)
            continue
        end

        % Check if it's a matrix or constant
        if ~contains(value, '[')
            v = str2double(value);
            if ~isnan(v) && isfinite(v)
                constNames{end+1} = name; %#ok<AGROW>
                constValues(end+1) = v;
            end
            continue
        end

        % Parse as matrix
        M = parseMatrixString(value);

        if isempty(M)
            continue
        end

        if size(M, 1) == 1
            tableNames{end+1} = name; %#ok<AGROW>
            tableValues{end+1} = M; %#ok<AGROW>
        else
            mapNames{end+1} = name; %#ok<AGROW>
            mapValues{end+1} = M; %#ok<AGROW>
        end
    end

    % Delete existing file
    if exist(outputFile, 'file')
        delete(outputFile);
    end

    % Write Constant sheet
    if ~isempty(constNames)
        constSheet = cell(numel(constNames) + 1, 2);
        constSheet{1,1} = 'Constant';
        constSheet{1,2} = 'Value';
        for i = 1:numel(constNames)
            constSheet{i+1,1} = constNames{i};
            constSheet{i+1,2} = roundToDecimal(constValues(i), decimals);
        end
        xlswrite(outputFile, constSheet, 'Constant');
    else
        xlswrite(outputFile, {'Constant','Value'}, 'Constant');
    end

    % Write Table sheet
    if ~isempty(tableNames)
        maxCols = 0;
        for i = 1:numel(tableNames)
            maxCols = max(maxCols, size(tableValues{i}, 2) + 1);
        end
        tableSheet = cell(numel(tableNames) + 1, maxCols);
        tableSheet{1,1} = 'Table';
        row = 2;
        for i = 1:numel(tableNames)
            tableSheet{row,1} = tableNames{i};
            vals = tableValues{i};
            for c = 1:size(vals, 2)
                tableSheet{row, c+1} = roundToDecimal(vals(1, c), decimals);
            end
            row = row + 1;
        end
        xlswrite(outputFile, tableSheet, 'Table');
    else
        xlswrite(outputFile, {'Table'}, 'Table');
    end

    % Write Map sheet
    if ~isempty(mapNames)
        % First pass: calculate total rows needed
        totalRows = 1; % Header row
        for i = 1:numel(mapNames)
            M = mapValues{i};
            totalRows = totalRows + size(M, 1) + 2; % data rows + blank row
        end
        
        % Pre-allocate with exact size needed
        maxCols = 0;
        for i = 1:numel(mapNames)
            maxCols = max(maxCols, size(mapValues{i}, 2) + 1);
        end
        
        mapSheet = cell(totalRows, maxCols);
        mapSheet{1,1} = 'Map';
        row = 2;
        
        for i = 1:numel(mapNames)
            M = mapValues{i};
            [r, c] = size(M);

            mapSheet{row,1} = mapNames{i};
            row = row + 1;

            for rr = 1:r
                for cc = 1:c
                    mapSheet{row + rr - 1, cc} = roundToDecimal(M(rr, cc), decimals);
                end
            end

            row = row + r + 1;
        end
        
        % Remove any trailing empty rows
        lastRow = 1;
        for r = size(mapSheet, 1):-1:1
            if ~all(cellfun(@isempty, mapSheet(r, :)))
                lastRow = r;
                break
            end
        end
        mapSheet = mapSheet(1:lastRow, :);
        
        xlswrite(outputFile, mapSheet, 'Map');
    else
        xlswrite(outputFile, {'Map'}, 'Map');
    end

    % Try to remove Sheet1 if it exists
    try
        if ispc
            excelApp = actxserver('Excel.Application');
            excelApp.Visible = false;
            excelBook = excelApp.Workbooks.Open(outputFile);
            
            % Find and delete Sheet1
            for i = 1:excelBook.Sheets.Count
                if strcmp(excelBook.Sheets.Item(i).Name, 'Sheet1')
                    excelBook.Sheets.Item(i).Delete();
                    break
                end
            end
            
            excelBook.Save();
            excelBook.Close();
            excelApp.Quit();
            delete(excelApp);
        end
    catch
        % If Sheet1 removal fails, just continue
    end

    fprintf('✓ Conversion complete: %s\n', outputFile);
    fprintf('  Constants: %d\n', numel(constNames));
    fprintf('  Tables (1D): %d\n', numel(tableNames));
    fprintf('  Maps (2D): %d\n', numel(mapNames));
end

function M = parseMatrixString(value)
    % Parse MATLAB matrix string including multi-line matrices
    value = strtrim(value);

    if isempty(value) || value(1) ~= '[' || value(end) ~= ']'
        M = [];
        return
    end

    % Remove brackets
    value = value(2:end-1);
    value = strrep(value, ',', ' ');
    value = strtrim(value);

    if isempty(value)
        M = [];
        return
    end

    rows = regexp(value, ';', 'split');
    M = [];

    for r = 1:numel(rows)
        rowStr = strtrim(rows{r});

        if isempty(rowStr)
            continue
        end

        nums = sscanf(rowStr, '%f');

        if isempty(nums)
            M = [];
            return
        end

        nums = nums(:).';

        if isempty(M)
            M = nums;
        else
            if numel(nums) ~= size(M, 2)
                M = [];
                return
            end
            M = [M; nums]; %#ok<AGROW>
        end
    end
end

function v = roundToDecimal(x, d)
    if ~isfinite(x)
        v = x;
        return
    end
    v = round(x * 10^d) / 10^d;
end
