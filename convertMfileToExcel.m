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

    % Read the m-file as text
    txt = fileread(inputFile);
    txt = strrep(txt, sprintf('\r\n'), sprintf('\n'));
    txt = strrep(txt, sprintf('\r'), sprintf('\n'));

    % Remove comments
    lines = regexp(txt, '\n', 'split');
    cleanTxt = '';
    for i = 1:numel(lines)
        line = lines{i};
        % Remove % comment
        idx = find(line == '%', 1);
        if ~isempty(idx)
            line = line(1:idx-1);
        end
        cleanTxt = [cleanTxt line newline]; %#ok<AGROW>
    end

    % Prepare containers
    constNames = {};
    constValues = [];
    tableNames = {};
    tableValues = {};
    mapNames = {};
    mapValues = {};

    % Extract complete statements: name = [...];
    % Must handle multi-line matrices
    expr = '([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(\[.*?\]);';
    tokens = regexp(cleanTxt, expr, 'tokens');

    for k = 1:numel(tokens)
        varName = strtrim(tokens{k}{1});
        varValue = strtrim(tokens{k}{2});

        if isempty(varName) || isempty(varValue)
            continue
        end

        % Classify: constant, table, or map
        if ~contains(varValue, '[')
            % Try to parse as a constant
            val = str2double(varValue);
            if ~isnan(val)
                constNames{end+1} = varName;
                constValues(end+1) = val;
            end
        else
            % Parse as matrix
            mat = parseMatrixString(varValue);
            
            if ~isempty(mat)
                if size(mat, 1) == 1
                    % 1D table (single row)
                    tableNames{end+1} = varName;
                    tableValues{end+1} = mat;
                else
                    % 2D map (multiple rows)
                    mapNames{end+1} = varName;
                    mapValues{end+1} = mat;
                end
            end
        end
    end

    % Delete file if it exists
    if exist(outputFile, 'file')
        delete(outputFile);
    end

    % Sheet 1: Constants
    constantData = {};
    constantData{1,1} = 'Constant';
    constantData{1,2} = 'Value';
    
    for k = 1:numel(constNames)
        constantData{k+1, 1} = constNames{k};
        constantData{k+1, 2} = roundToDecimal(constValues(k), decimals);
    end

    % Sheet 2: Tables (1D arrays)
    tableData = {};
    tableData{1,1} = 'Table';
    row = 2;
    
    for k = 1:numel(tableNames)
        tableData{row, 1} = tableNames{k};
        vals = tableValues{k};
        for c = 1:size(vals, 2)
            tableData{row, c+1} = roundToDecimal(vals(1, c), decimals);
        end
        row = row + 1;
    end

    % Sheet 3: Maps (2D arrays)
    mapData = {};
    mapData{1,1} = 'Map';
    mapData{1,2} = 'Dimensions';
    row = 2;
    
    for k = 1:numel(mapNames)
        M = mapValues{k};
        [rows, cols] = size(M);
        
        % Write map name and dimensions
        mapData{row, 1} = mapNames{k};
        mapData{row, 2} = sprintf('%d rows x %d cols', rows, cols);
        row = row + 1;
        
        % Write the matrix: each row on a new row in Excel
        for i = 1:rows
            for j = 1:cols
                mapData{row, j} = roundToDecimal(M(i, j), decimals);
            end
            row = row + 1;
        end
        
        row = row + 1; % Add blank row between maps
    end

    % Write to Excel using xlswrite
    xlswrite(outputFile, constantData, 'Constant');
    xlswrite(outputFile, tableData, 'Table');
    xlswrite(outputFile, mapData, 'Map');

    fprintf('✓ Conversion complete: %s\n', outputFile);
    fprintf('  Constants: %d\n', numel(constNames));
    fprintf('  Tables (1D): %d\n', numel(tableNames));
    fprintf('  Maps (2D): %d\n', numel(mapNames));
end

function val = roundToDecimal(v, d)
    % Round to d decimal places
    if isfinite(v)
        val = round(v * 10^d) / 10^d;
    else
        val = v;
    end
end

function M = parseMatrixString(s)
    % Parse MATLAB matrix string with support for:
    % [1 2 3 4]
    % [1,2,3,4]
    % [1 2 3; 4 5 6]
    % [1,2,3;
    %  4,5,6]
    
    s = strtrim(s);
    
    if isempty(s) || s(1) ~= '[' || s(end) ~= ']'
        M = [];
        return
    end

    % Remove brackets
    s = s(2:end-1);
    s = strtrim(s);

    if isempty(s)
        M = [];
        return
    end

    % Normalize: replace commas with spaces
    s = strrep(s, ',', ' ');
    
    % Check if there are semicolons (row separators)
    if contains(s, ';')
        % Multi-row matrix: split by semicolon
        rowStrs = regexp(s, ';', 'split');
        M = [];
        
        for r = 1:numel(rowStrs)
            rowStr = strtrim(rowStrs{r});
            
            if isempty(rowStr)
                continue
            end

            % Parse numbers in this row using sscanf
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
    else
        % Single row matrix: no semicolons
        nums = sscanf(s, '%f');
        
        if isempty(nums)
            M = [];
            return
        end

        M = nums(:).'; % Ensure row vector
    end
end
