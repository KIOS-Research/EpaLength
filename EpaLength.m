function EpaLength(units)
    % units  - meters or feet
    % Create table of diameters for multiple Input files 
    addpath(genpath(pwd));
    
    filename = [units,'_data.xls'];
    try
    xlswrite(filename,{''},'Sheet1','A1');  
    catch 
        errormsg = strcat('The file "', filename, '" is open. please close it.');
        warning on;
        warning(errormsg);
        return;
    end
    
    dir_struct = dir(strcat([pwd,'\networks\'],'*'));
    [sorted_names,~] = sortrows({dir_struct.name}');
    DMAs = sorted_names(3:end);
    format long g;
    t=1;
    for i=1:length(DMAs)
        DMAstr=DMAs{i};
        d=epanet(['networks/',DMAstr]);
        if ~strcmp(d.LinkLengthsUnits,units)
            if strcmp(units,'feet')
                d.setBinFlowUnitsGPM;
            elseif strcmp(units,'meters')
                d.setBinFlowUnitsLPS;
            end
        end
        parameters=[];
        pipesID={};
        linklengths=d.getLinkLength;
        linkdiameters=d.getLinkDiameter;
        parameters(:,1)=(linklengths(1:d.LinkPipeCount));% length 
        parameters(:,2)=(linkdiameters(1:d.LinkPipeCount));% diameter
        DMAparameters{t}=parameters;
        fdiameters{t}=unique(DMAparameters{t}(:,2));
        for m=1:length(fdiameters{t})
            flengths{t}(m)=sum(DMAparameters{t}(find(DMAparameters{t}(:,2)==fdiameters{t}(m)),1));
        end
         fparameters{t}=[fdiameters{t}'; flengths{t}];
        t=t+1;
        d.unload;
    end
    m=1;
    for u=1:length(DMAparameters)
        for k=1:length(DMAparameters{u}(:,2)) 
            alldiameters(m)=DMAparameters{u}(k,2) ;
            m=m+1;
        end
    end
    DIAMETERS_ALL_MM=sort(unique(alldiameters),'descend');
    for t=1:length(fparameters)
        for pp=1:length(DIAMETERS_ALL_MM)
            index=find(fparameters{t}(1,:)==DIAMETERS_ALL_MM(pp));
            if isempty(index)
                finalTABLE(t,pp)=0;
            else
                finalTABLE(t,pp)=fparameters{t}(2,index);
            end
        end
    end
    a=size(finalTABLE);
    res = zeros(a(1)+1,a(2)+1);
    res(2:end,1:end-1)=finalTABLE;
    res(1,1:end-1)=DIAMETERS_ALL_MM;
    resTitle = {'DMA', DMAs{:}}';
    b = col(length(res));
    for i=1:length(resTitle)
       res(i,end) = sum(res(i,:));
    end
    Totals = res(1:end,end);
    TotalsRes = sum(Totals(2:end));
        
    xlswrite(filename,resTitle,'Sheet1','A2');  
    xlswrite(filename,{['Total Lengths DMAs(',units,')']},'Sheet1',b);     
    xlswrite(filename,res,'Sheet1','B2');     
    xlswrite(filename,{''},'Sheet1',[b,'2']);     
    xlswrite(filename,{TotalsRes},'Sheet1',[b,num2str(length(resTitle)+3)]);      
    bb = col(length(res)-1);
    xlswrite(filename,{'Total Length'},'Sheet1',[bb,num2str(length(resTitle)+3)]);     

    ex = actxserver('excel.application');    
    ex.Workbooks.Open([pwd,'\',filename]);
    ex.Columns.Item(b).ColumnWidth = 25;
    ex.Visible = 1;
    ex.Columns.Item(bb).ColumnWidth = 12;
    ex.Columns.Item('A').ColumnWidth = 22;
    ex.Selection.WrapText = true;
    fprintf('Run was Successful.\n\n') 
   
    
function b = col(res)   
    mm=26;
    num=res+1;
    n = ceil(log(num)/log(mm));   
    d = cumsum(mm.^(0:n+1));    
    n = find(num >= d, 1, 'last');  
    d = d(n:-1:1);                
    r = mod(floor((num-d)./mm.^(n-1:-1:0)), mm) + 1;   
    b = char(r+64); 
