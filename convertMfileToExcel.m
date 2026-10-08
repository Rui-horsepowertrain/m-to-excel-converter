function convertMfileToExcel(inputFile, outputFile, decimals)
    % convertMfileToExcel Convert MATLAB .m file constants, tables, and maps to Excel
    %
    % Example:
    %   convertMfileToExcel('input.m', 'output.xlsx', 3);

    if nargin < 3 || isempty(decimals)
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

    % Match assignments like: name = value;
    expr = '([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(\[[\s\S]*?\]|[^;\n]*);';
    toks = regexp(cleanTxt, expr, 'tokens');

    for k = 1:numel(toks)
        name = toks{k}{1};
        value = strtrim(toks{k}{2});

        if isempty(value)
            continue
        end

        % Constant
        if ~contains(value, '[')
            val = str2double(value);
            if ~isnan(val) && isfinite(val)
                constNames{end+1} = name; %#ok<AGROW>
                constValues(end+1) = val;
            end
            continue
        end

        % Matrix
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

    if exist(outputFile, 'file')
        delete(outputFile);
    end

    % Create a temporary file first
    tempFile = [tempname '.xlsx'];

    % ------------------------
    % Constant sheet
    % ------------------------
    constSheet = {'Constant', 'Value'};
    for i = 1:numel(constNames)
        constSheet{i+1,1} = constNames{i};
        constSheet{i+1,2} = roundToDecimal(constValues(i), decimals);
    end
    xlswrite(tempFile, constSheet, 'Constant');

    % ------------------------
    % Table sheet
    % ------------------------
    if isempty(tableNames)
        xlswrite(tempFile, {'Table'}, 'Table');
    else
        tableSheet = {'Table'};
        row = 2;
        for i = 1:numel(tableNames)
            tableSheet{row,1} = tableNames{i};
            vals = tableValues{i};
            for c = 1:numel(vals)
                tableSheet{row, c+1} = roundToDecimal(vals(c), decimals);
            end
            row = row + 1;
        end
        xlswrite(tempFile, tableSheet, 'Table');
    end

    % ------------------------
    % Map sheet
    % ------------------------
    if isempty(mapNames)
        xlswrite(tempFile, {'Map'}, 'Map');
    else
        mapSheet = {'Map'};
        row = 2;
        for i = 1:numel(mapNames)
            M = mapValues{i};
            mapSheet{row,1} = mapNames{i};
            row = row + 1;

            for r = 1:size(M,1)
                for c = 1:size(M,2)
                    mapSheet{row + r - 1, c} = roundToDecimal(M(r,c), decimals);
                end
            end

            row = row + size(M,1) + 1;
        end

        xlswrite(tempFile, mapSheet, 'Map');
    end

    % Remove Sheet1 and rename temp file
    removeSheet1AndFinalize(tempFile, outputFile);

    fprintf('✓ Done: %s\n', outputFile);
    fprintf('  Constants: %d | Tables: %d | Maps: %d\n', ...
        numel(constNames), numel(tableNames), numel(mapNames));
end

function removeSheet1AndFinalize(inputFile, outputFile)
    % Use Excel COM to remove Sheet1 and save final file
    
    if ispc
        try
            % Open Excel
            excelApp = actxserver('Excel.Application');
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            
            % Open workbook
            workbook = excelApp.Workbooks.Open(inputFile);
            
            % Get all sheet names
            sheetCount = workbook.Sheets.Count;
            sheet1Found = false;
            
            % Delete Sheet1 if it exists
            for i = 1:sheetCount
                sheetName = workbook.Sheets.Item(i).Name;
                if strcmp(sheetName, 'Sheet1')
                    workbook.Sheets.Item(i).Delete();
                    sheet1Found = true;
                    break;
                end
            end
            
            % Save and close
            workbook.SaveAs(outputFile);
            workbook.Close();
            excelApp.Quit();
            delete(excelApp);
            
            % Clean up temp file
            if exist(inputFile, 'file')
                delete(inputFile);
            end
        catch ME
            % If Excel fails, copy temp file and try alternative method
            copyfile(inputFile, outputFile);
            if exist(inputFile, 'file')
                delete(inputFile);
            end
            warning('Could not delete Sheet1 via Excel COM. Attempting alternative method.');
            
            % Try using Python if available (fallback)
            try
                py_deleteSheet1(outputFile);
            catch
                warning('Could not remove Sheet1. Please delete manually.');
            end
        end
    else
        % Non-Windows system - just copy and warn
        copyfile(inputFile, outputFile);
        if exist(inputFile, 'file')
            delete(inputFile);
        end
        warning('Sheet1 removal requires Windows with Excel installed.');
    end
end

function py_deleteSheet1(excelFile)
    % Try to remove Sheet1 using Python (openpyxl)
    try
        py.sys.path.insert(int64(0), pwd());
        pyCmd = sprintf(['import openpyxl; ' ...
                        'wb = openpyxl.load_workbook("%s"); ' ...
                        'if "Sheet1" in wb.sheetnames: wb.remove(wb["Sheet1"]); ' ...
                        'wb.save("%s")'], excelFile, excelFile);
        py.eval(pyCmd);
    catch
        % Python method also failed, ignore
    end
end

function M = parseMatrixString(s)
    s = strtrim(s);

    if isempty(s) || s(1) ~= '[' || s(end) ~= ']'
        M = [];
        return
    end

    s = s(2:end-1);
    s = strrep(s, ',', ' ');
    s = strtrim(s);

    if isempty(s)
        M = [];
        return
    end

    rows = regexp(s, ';', 'split');

    if numel(rows) == 1
        nums = sscanf(rows{1}, '%f');
        M = reshape(nums, 1, numel(nums));
        return
    end

    M = [];
    for i = 1:numel(rows)
        rowStr = strtrim(rows{i});
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
            if numel(nums) ~= size(M,2)
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
