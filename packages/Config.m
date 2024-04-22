 classdef Config
    methods(Static)        
        function v = get_value(cell)
            if iscell(cell)
                cell = cell{1};
            end
            if isnumeric(cell)
                v = cell;
            elseif ischar(cell)
                if strcmpi(cell, 'NaN')
                    v = NaN;
                else
                    v = cell;
                end
            else
                v = NaN;
            end   
        end
        
        function v = get_number(cell)
            if iscell(cell)
                cell = cell{1};
            end
            if isnumeric(cell)
                v = cell;
            elseif cell == 'NaN'
                v = NaN;
            end
            if isnan(v)
                v = 0;
            end
        end
        
        function [trackX, trackY, trackZ, resultCurvature] = get_track(dst, targetCurvature, roadGradient, desiredSpeed)
            initialStraight = 20;
            curvatureGradient = 0.0005;
            distanceGradientRampEnd = initialStraight + ceil(desiredSpeed / 3.6 * 5 / 10) * 10;
            
            trackX = zeros(1, length(dst));
            trackY = zeros(1, length(dst));
            trackZ = zeros(1, length(dst));
            phi = zeros(1, length(dst));
            resultCurvature = zeros(1, length(dst));
            
            curvature = 0;
            for i = 2:length(dst)        
                if dst(i) > initialStraight
                    if dst(i) < distanceGradientRampEnd
                        curvature = (1-cos((dst(i) - initialStraight)/(distanceGradientRampEnd - initialStraight) * pi))/2 * targetCurvature;
                    else
                        curvature = targetCurvature;
                    end
                    %if curvature < targetCurvature
                    %    curvature = min(curvature + curvatureGradient, targetCurvature);                        
                    %end
                end
                resultCurvature(i) = curvature;
                deltaDst = dst(i) - dst(i-1);
                deltaPhi = deltaDst * curvature;
                if curvature == 0
                    x1 = deltaDst;
                    y1 = 0;
                else
                    radius = 1 / curvature;
                    x1 = radius * sin(deltaPhi);
                    y1 = radius * (1 - cos(deltaPhi));
                end
                
                trackX(i) = trackX(i-1) + x1 * cos(phi(i-1)) - y1 * sin(phi(i-1));
                trackY(i) = trackY(i-1) + y1 * cos(phi(i-1)) + x1 * sin(phi(i-1));
                trackZ(i) = trackZ(i-1) + roadGradient / 100 * deltaDst;
                phi(i) = phi(i - 1) + deltaPhi;  
            end
            
            if curvature ~= targetCurvature
                error("The specified curvature (%s m) could not be reached before the end of the track (%s m) with a curve gradient of %s 1/m", num2str(targetCurvature), num2str(dst(end)), num2str(curvatureGradient))
            end
        end        
        
        function xlcol_addr=num2xlcol(col_num)
            n=1;
            while col_num>26*(26^n-1)/25
                n=n+1;
            end
            base_26=zeros(1,n);
            tmp_var=-1+col_num-26*(26^(n-1)-1)/25;
            for k=1:n
                divisor=26^(n-k);
                remainder=mod(tmp_var,divisor);
                base_26(k)=65+(tmp_var-remainder)/divisor;
                tmp_var=remainder;
            end
            xlcol_addr=char(base_26);
        end
        
        function xlcol_num=xlcol2num(xlcol_addr)
            if ischar(xlcol_addr) && ~any(~isstrprop(xlcol_addr,"upper"))
                xlcol_num=0;
                n=length(xlcol_addr);
                for k=1:n
                    xlcol_num=xlcol_num+(double(xlcol_addr(k)-64))*26^(n-k);
                end
            else
                error('Not a valid character')
            end
        end
        
        function value = readConfig(configLines, section, key)
            key = lower(strtrim(key));
            inSection = 0;
            for iLine = 1 : length(configLines)
                configLine = strtrim(configLines(iLine));
                if ~inSection
                    if strcmpi(configLine, ['[', section, ']'])
                        inSection = 1;
                    end
                else
                    if startsWith(configLine, '[')
                        error('Section %s not found in config.ini', section)
                    end
                    if startsWith(lower(configLine), key)
                        if count(configLine, '=') ~= 1
                            error('Invalid line (%s) in config.ini', configLine)
                        else
                            splitLine = split(configLine, '=');
                            value = strtrim(splitLine(2));
                            return
                        end
                    end
                end

            end
            error('Key %s not found in section %s in config.ini', key, section)
        end
        
        function value = readConfigNumber(configLines, section, key)
            value = str2num(Config.readConfig(configLines, section, key));
        end
    end
 end