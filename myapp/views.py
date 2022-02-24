from django.shortcuts import render
import pandas as pd
import xlsxwriter
from io import BytesIO
from django.http.response import HttpResponse
import datetime
from tzlocal import get_localzone


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

def excel_download(request):
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

        timestr = datetime.datetime.now().astimezone(get_localzone()).strftime('%Y%m%d-%H%M%S')

        sheet_name = 'Sheet1'
        with BytesIO() as b:
            # Use the StringIO object as the filehandle.
            writer = pd.ExcelWriter(b, engine='xlsxwriter',
                                    datetime_format='mmm-dd-yyyy',
                                    date_format='mmmm-dd-yyyy')
            output.to_excel(writer, sheet_name=sheet_name)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            # Get the dimensions of the dataframe.
            (max_row, max_col) = output.shape
            # Set the column widths, to make the dates clearer.
            worksheet.set_column(0, max_col+10, 20)
            writer.save()
            filename = 'sample_collection_schedule_generated_' + timestr
            content_type = 'application/vnd.ms-excel'
            response = HttpResponse(b.getvalue(), content_type=content_type)
            response['Content-Disposition'] = 'attachment; filename="' + filename + '.xlsx"'
            return response










