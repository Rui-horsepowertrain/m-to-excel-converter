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

    % Prepare containers
    constNames = {};
    constValues = [];
    tableNames = {};
    tableValues = {};
    mapNames = {};
    mapValues = {};

    % Split into statements (separated by semicolon at end of line)
    % But be careful: semicolons can appear inside matrix definitions
    
    % First, remove comments
    lines = regexp(txt, '\n', 'split');
    cleanTxt = '';
    for i = 1:numel(lines)
        line = lines{i};
        % Remove % comment
        idx = find(line == '%', 1);
        if ~isempty(idx)
            line = line(1:idx-1);
        end
        cleanTxt = [cleanTxt line '\n']; %#ok<AGROW>
    end

    % Now extract statements: look for pattern: name = [...];
    % Use regex to find complete assignments
    
    % Pattern: variable_name = ... ;
    pattern = '(\w+)\s*=\s*([^;]*;)';
    tokens = regexp(cleanTxt, pattern, 'tokens');

    for k = 1:numel(tokens)
        varName = strtrim(tokens{k}{1});
        varValue = strtrim(tokens{k}{2});
        
        % Remove trailing semicolon
        varValue = regexprep(varValue, ';\s*$', '');
        varValue = strtrim(varValue);

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

    % Build Excel sheets
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
    row = 2;
    
    for k = 1:numel(mapNames)
        mapData{row, 1} = mapNames{k};
        M = mapValues{k};
        [rows, cols] = size(M);
        
        % Write map name and dimensions in first row
        mapData{row, 2} = sprintf('%dx%d', rows, cols);
        row = row + 1;
        
        % Write the matrix rows
        for i = 1:rows
            for j = 1:cols
                mapData{row, j+1} = roundToDecimal(M(i, j), decimals);
            end
            row = row + 1;
        end
        
        row = row + 1; % Add blank row between maps
    end

    % Write to Excel
    writetable(cell2table(constantData), outputFile, 'Sheet', 'Constant', 'WriteVariableNames', false);
    writetable(cell2table(tableData), outputFile, 'Sheet', 'Table', 'WriteVariableNames', false);
    writetable(cell2table(mapData), outputFile, 'Sheet', 'Map', 'WriteVariableNames', false);

    fprintf('✓ Conversion complete: %s\n', outputFile);
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
    % Parse MATLAB matrix string: [1 2 3; 4 5 6] or [1,2,3; 4,5,6]
    % Handles multi-line input
    
    s = strtrim(s);
    
    if isempty(s) || ~startsWith(s, '[') || ~endsWith(s, ']')
        M = [];
        return
    end

    % Remove outer brackets
    s = s(2:end-1);
    s = strtrim(s);

    % Normalize separators:
    % - Semicolon (;) = row separator
    % - Space or comma = column separator
    
    % Replace commas with spaces
    s = strrep(s, ',', ' ');
    
    % Normalize whitespace: multiple spaces/newlines → single space
    s = regexprep(s, '\s+', ' ');
    s = strtrim(s);

    % Split into rows by semicolon
    rowStrs = strsplit(s, ';');
    
    M = [];
    
    for r = 1:numel(rowStrs)
        rowStr = strtrim(rowStrs{r});
        
        if isempty(rowStr)
            continue
        end

        % Parse numbers in this row
        nums = str2num(rowStr); %#ok<ST2NM>
        
        if isempty(nums)
            % Failed to parse
            M = [];
            return
        end

        % Ensure nums is a row vector
        if iscolumn(nums)
            nums = nums.';
        end

        % Add to matrix
        if isempty(M)
            M = nums;
        else
            % Check column count matches
            if size(nums, 2) ~= size(M, 2)
                % Dimension mismatch
                M = [];
                return
            end
            M = [M; nums]; %#ok<AGROW>
        end
    end
end

function parts = strsplit(s, delim)
    % Simple string split by delimiter
    parts = regexp(s, sprintf('[^%s]+', regexptranslate('escape', delim)), 'match');
end
