from django.shortcuts import render
import pandas as pd



def index(request):
    if "GET" == request.method:
        return render(request, 'myapp/index.html', {})
    else:
        excel_file = request.FILES["excel_file"]
        dfs = pd.read_excel(excel_file, sheet_name='Sheet1')
        field_names = ["First Collection", "Second Collection", "Third Collection"]
        output = pd.DataFrame()
        for field in field_names:
            group_val = pd.DataFrame(dfs.groupby([field, "Farm", "House"]).size())
            output = pd.concat([output, group_val])
        output.columns = ['Count']
        ret_html = output.to_html()

        return render(request, 'myapp/index.html', {"excel_data":ret_html})









