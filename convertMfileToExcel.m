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
            % Try to parse as constant
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
            % 1D table
            tableNames{end+1} = name; %#ok<AGROW>
            tableValues{end+1} = M; %#ok<AGROW>
        else
            % 2D map
            mapNames{end+1} = name; %#ok<AGROW>
            mapValues{end+1} = M; %#ok<AGROW>
        end
    end

    % Delete existing file
    if exist(outputFile, 'file')
        delete(outputFile);
    end

    % Create empty workbook first (this creates Sheet1 by default)
    xlswrite(outputFile, {''}, 'Constant');

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
    end

    % Write Table sheet
    if ~isempty(tableNames)
        tableSheet = cell(numel(tableNames) + 1, 20);
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
    end

    % Write Map sheet
    if ~isempty(mapNames)
        mapSheet = cell(100, 20); % Pre-allocate
        mapSheet{1,1} = 'Map';
        mapSheet{1,2} = 'Dimensions';
        row = 2;
        for i = 1:numel(mapNames)
            M = mapValues{i};
            [r, c] = size(M);

            mapSheet{row,1} = mapNames{i};
            mapSheet{row,2} = sprintf('%d x %d', r, c);
            row = row + 1;

            for rr = 1:r
                for cc = 1:c
                    mapSheet{row + rr - 1, cc + 1} = roundToDecimal(M(rr, cc), decimals);
                end
            end

            row = row + r + 1;
        end
        % Trim empty rows
        mapSheet(row:end, :) = [];
        xlswrite(outputFile, mapSheet, 'Map');
    end

    % Delete default Sheet1 if it exists and we created other sheets
    try
        excelApp = actxserver('Excel.Application');
        excelApp.Visible = false;
        excelBook = excelApp.Workbooks.Open(outputFile);
        
        sheetNames = {};
        for i = 1:excelBook.Sheets.Count
            sheetNames{i} = excelBook.Sheets.Item(i).Name;
        end
        
        % Delete Sheet1 if it exists and is not the only sheet
        sheet1Idx = find(strcmp(sheetNames, 'Sheet1'), 1);
        if ~isempty(sheet1Idx) && excelBook.Sheets.Count > 1
            excelBook.Sheets.Item(sheet1Idx).Delete();
        end
        
        excelBook.Save();
        excelBook.Close();
        excelApp.Quit();
        delete(excelApp);
    catch
        % If ActiveX fails, just warn user
        fprintf('Warning: Could not remove default Sheet1. Please delete manually.\n');
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
    
    % Replace commas with spaces
    value = strrep(value, ',', ' ');
    value = strtrim(value);

    if isempty(value)
        M = [];
        return
    end

    % Split by semicolon (row separator)
    rows = regexp(value, ';', 'split');
    M = [];

    for r = 1:numel(rows)
        rowStr = strtrim(rows{r});
        
        if isempty(rowStr)
            continue
        end

        % Parse numbers in this row
        nums = sscanf(rowStr, '%f');
        
        if isempty(nums)
            M = [];
            return
        end

        nums = nums(:).'; % Ensure row vector

        % Add to matrix
        if isempty(M)
            M = nums;
        else
            % Check column count matches
            if numel(nums) ~= size(M, 2)
                M = [];
                return
            end
            M = [M; nums]; %#ok<AGROW>
        end
    end
end

function v = roundToDecimal(x, d)
    % Round to d decimal places
    
    if ~isfinite(x)
        v = x;
        return
    end
    v = round(x * 10^d) / 10^d;
end
