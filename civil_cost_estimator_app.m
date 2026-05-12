function civil_cost_estimator_app
    fig = uifigure('Name','Civil Cost Estimator','Position',[100 100 1080 680]);
    fig.Color = [0.97 0.97 0.98];

    gl = uigridlayout(fig,[1 2]);
    gl.ColumnWidth = {380,'1x'};
    gl.Padding = [12 12 12 12];
    gl.ColumnSpacing = 12;

    left = uipanel(gl,'Title','User Input');
    right = uipanel(gl,'Title','Results');

    lgl = uigridlayout(left,[15 2]);
    lgl.RowHeight = repmat({28},1,15);
    lgl.ColumnWidth = {'1x','1x'};
    lgl.Padding = [10 10 10 10];
    lgl.RowSpacing = 8;

    uilabel(lgl,'Text','Material cost');
    mat = uieditfield(lgl,'numeric','AllowEmpty','on','Value',[],'Placeholder','Enter material cost');

    uilabel(lgl,'Text','Labor cost');
    lab = uieditfield(lgl,'numeric','AllowEmpty','on','Value',[],'Placeholder','Enter labor cost');

    uilabel(lgl,'Text','Equipment cost');
    eqp = uieditfield(lgl,'numeric','AllowEmpty','on','Value',[],'Placeholder','Enter equipment cost');

    uilabel(lgl,'Text','Subcontract cost');
    sub = uieditfield(lgl,'numeric','AllowEmpty','on','Value',[],'Placeholder','Enter subcontract cost');

    uilabel(lgl,'Text','Overhead %');
    ovh = uieditfield(lgl,'numeric','AllowEmpty','on','Value',[],'Placeholder','Enter overhead %');

    uilabel(lgl,'Text','Contingency %');
    con = uieditfield(lgl,'numeric','AllowEmpty','on','Value',[],'Placeholder','Enter contingency %');

    uilabel(lgl,'Text','Profit %');
    prof = uieditfield(lgl,'numeric','AllowEmpty','on','Value',[],'Placeholder','Enter profit %');

    uilabel(lgl,'Text','Tax %');
    tax = uieditfield(lgl,'numeric','AllowEmpty','on','Value',[],'Placeholder','Enter tax %');

    btn = uibutton(lgl,'Text','Calculate','FontWeight','bold','ButtonPushedFcn',@calculate);
    btn.Layout.Row = 9;
    btn.Layout.Column = [1 2];

    resetbtn = uibutton(lgl,'Text','Reset','ButtonPushedFcn',@resetFields);
    resetbtn.Layout.Row = 10;
    resetbtn.Layout.Column = [1 2];

    exportbtn = uibutton(lgl,'Text','Export CSV','ButtonPushedFcn',@exportCSV);
    exportbtn.Layout.Row = 11;
    exportbtn.Layout.Column = [1 2];

    info = uilabel(lgl,'Text','Fill all inputs then press Calculate.','FontAngle','italic');
    info.Layout.Row = 12;
    info.Layout.Column = [1 2];

    warn = uilabel(lgl,'Text','', ...
        'FontWeight','bold', ...
        'FontColor',[0.85 0.2 0.2]);
    warn.Layout.Row = 13;
    warn.Layout.Column = [1 2];

    help1 = uilabel(lgl,'Text','Made for civil engineering portfolio.');
    help1.Layout.Row = 14;
    help1.Layout.Column = [1 2];

    help2 = uilabel(lgl,'Text','Dynamic user input + chart + Excel export.');
    help2.Layout.Row = 15;
    help2.Layout.Column = [1 2];

    rgl = uigridlayout(right,[3 1]);
    rgl.RowHeight = {100,250,'1x'};
    rgl.Padding = [10 10 10 10];

    summary = uipanel(rgl,'Title','Summary');
    sgl = uigridlayout(summary,[1 4]);
    sgl.ColumnWidth = {'1x','1x','1x','1x'};
    sgl.Padding = [8 8 8 8];

    totalBox = uilabel(sgl,'Text','$0','FontSize',22,'FontWeight','bold','HorizontalAlignment','center');
    directBox = uilabel(sgl,'Text','Direct\n$0','HorizontalAlignment','center');
    subtotalBox = uilabel(sgl,'Text','Subtotal\n$0','HorizontalAlignment','center');
    markupBox = uilabel(sgl,'Text','Markup\n0%','HorizontalAlignment','center');

    tablePanel = uipanel(rgl,'Title','Cost Breakdown Table');
    tgl = uigridlayout(tablePanel,[1 1]);
    tgl.Padding = [8 8 8 8];
    tbl = uitable(tgl,'ColumnName',{'Category','Amount','Percent of Final'},'ColumnEditable',[false false false]);

    ax = uiaxes(rgl);
    title(ax,'Cost Breakdown');
    ylabel(ax,'Amount');

    function v = getValue(field, fieldName)
        v = field.Value;
        if isempty(v) || isnan(v)
            error('%s is required.', fieldName);
        end
    end

    function [data, direct, overhead, contingency, profitAmt, taxAmt, subtotal, final] = compute()
        material = getValue(mat,'Material cost');
        labor = getValue(lab,'Labor cost');
        equipment = getValue(eqp,'Equipment cost');
        subcontract = getValue(sub,'Subcontract cost');
        overheadPct = getValue(ovh,'Overhead %');
        contingencyPct = getValue(con,'Contingency %');
        profitPct = getValue(prof,'Profit %');
        taxPct = getValue(tax,'Tax %');

        direct = material + labor + equipment + subcontract;
        overhead = direct * overheadPct/100;
        contingency = direct * contingencyPct/100;
        profitAmt = (direct + overhead + contingency) * profitPct/100;
        subtotal = direct + overhead + contingency + profitAmt;
        taxAmt = subtotal * taxPct/100;
        final = subtotal + taxAmt;

        data = {
            'Material', material, 100*material/max(final,eps);
            'Labor', labor, 100*labor/max(final,eps);
            'Equipment', equipment, 100*equipment/max(final,eps);
            'Subcontract', subcontract, 100*subcontract/max(final,eps);
            'Overhead', overhead, 100*overhead/max(final,eps);
            'Contingency', contingency, 100*contingency/max(final,eps);
            'Profit', profitAmt, 100*profitAmt/max(final,eps);
            'Tax', taxAmt, 100*taxAmt/max(final,eps)
        };
    end

    function saveToExcel(material,labor,equipment,subcontract,overheadPct,contingencyPct,profitPct,taxPct,direct,overhead,contingency,profitAmt,subtotal,taxAmt,final)
        fileName = 'civil_cost_estimator_results.xlsx';
        newRow = table( ...
            datetime('now'), material, labor, equipment, subcontract, overheadPct, contingencyPct, profitPct, taxPct, ...
            direct, overhead, contingency, profitAmt, subtotal, taxAmt, final, ...
            'VariableNames', {'Timestamp','MaterialCost','LaborCost','EquipmentCost','SubcontractCost','OverheadPct','ContingencyPct','ProfitPct','TaxPct','DirectCost','OverheadCost','ContingencyCost','ProfitCost','Subtotal','TaxAmount','FinalCost'} );

        if isfile(fileName)
            oldData = readtable(fileName);
            allData = [oldData; newRow];
        else
            allData = newRow;
        end

        writetable(allData,fileName);
    end

    function calculate(~,~)
        try
            warn.Text = '';
            material = getValue(mat,'Material cost');
            labor = getValue(lab,'Labor cost');
            equipment = getValue(eqp,'Equipment cost');
            subcontract = getValue(sub,'Subcontract cost');
            overheadPct = getValue(ovh,'Overhead %');
            contingencyPct = getValue(con,'Contingency %');
            profitPct = getValue(prof,'Profit %');
            taxPct = getValue(tax,'Tax %');

            [data, direct, overhead, contingency, profitAmt, taxAmt, subtotal, final] = compute();
            tbl.Data = data;
            tbl.ColumnFormat = {'char','bank','bank'};

            totalBox.Text = sprintf('$%.2f', final);
            directBox.Text = sprintf('Direct\n$%.2f', direct);
            subtotalBox.Text = sprintf('Subtotal\n$%.2f', subtotal);
            markupBox.Text = sprintf('Markup\n%.1f%%', 100*(final-direct)/max(direct,eps));

            cla(ax);
            vals = cell2mat(data(:,2));
            bar(ax, vals);
            ax.XTick = 1:numel(vals);
            ax.XTickLabel = data(:,1);
            ax.XTickLabelRotation = 25;
            ylabel(ax,'Amount');
            title(ax,'Cost Breakdown');
            grid(ax,'on');

            saveToExcel(material,labor,equipment,subcontract,overheadPct,contingencyPct,profitPct,taxPct,direct,overhead,contingency,profitAmt,subtotal,taxAmt,final);
            uialert(fig,'Saved to civil_cost_estimator_results.xlsx','Saved');
        catch ME
            warn.Text = ME.message;
        end
    end

    function resetFields(~,~)
        mat.Value = [];
        lab.Value = [];
        eqp.Value = [];
        sub.Value = [];
        ovh.Value = [];
        con.Value = [];
        prof.Value = [];
        tax.Value = [];
        tbl.Data = {};
        cla(ax);
        totalBox.Text = '$0';
        directBox.Text = 'Direct\n$0';
        subtotalBox.Text = 'Subtotal\n$0';
        markupBox.Text = 'Markup\n0%';
        warn.Text = '';
    end

    function exportCSV(~,~)
        try
            [data, direct, overhead, contingency, profitAmt, taxAmt, subtotal, final] = compute();
            T = cell2table(data,'VariableNames',{'Category','Amount','PercentOfFinal'});
            T.Properties.Description = sprintf('Direct=%.2f; Overhead=%.2f; Contingency=%.2f; Profit=%.2f; Tax=%.2f; Subtotal=%.2f; Final=%.2f', direct, overhead, contingency, profitAmt, taxAmt, subtotal, final);
            writetable(T,'civil_cost_breakdown.csv');
            uialert(fig,'CSV exported as civil_cost_breakdown.csv in the current folder.','Export complete');
        catch ME
            uialert(fig,ME.message,'Export failed');
        end
    end
end