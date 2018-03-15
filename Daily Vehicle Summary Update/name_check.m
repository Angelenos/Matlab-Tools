function  n = name_check()
    saveLoc = 'C:\work\OBD reliability\Issue Tracking\Vehicle Status Summary\';
    file_all = dir(saveLoc);
    n = 0;
    
    for i = 1 : length(file_all)
        file_name = file_all(i).name;
        if(length(strfind(file_name,'csv')) > 1)
            file_date = regexp(file_name,'\d{4}');
            file_date = file_name(file_date(1) : file_date(1)+3);
            if(~isempty(strfind(file_name,'Software Detail')))
                movefile (strcat(saveLoc, file_name), strcat(...
                    saveLoc, 'CONTROLTEC Vehicle Software Detail_', ...
                    file_date,'.csv'));
            elseif(~isempty(strfind(file_name,'Detail')))
                movefile (strcat(saveLoc, file_name), strcat(...
                    saveLoc, 'CONTROLTEC Vehicle Detail_', ...
                    file_date,'.csv'));
            else
                movefile (strcat(saveLoc, file_name), strcat(...
                    saveLoc, 'CONTROLTEC Vehicle Properties_', ...
                    file_date,'.csv'));
            end
             n = n + 1;
        end
    end
end