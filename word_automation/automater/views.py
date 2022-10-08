from asyncio.windows_events import NULL
from csv import excel_tab
import datetime
import random
from django.http import HttpResponse
from django.shortcuts import render
import pandas as pd
from.models import excel
from docxtpl import DocxTemplate
import pandas as pd
import random
import numpy as np
import jinja2
import datetime
from io import StringIO
import zipfile
import os, io
import glob

# Create your views here.
def home(request):
    if request.method=='POST' and request.FILES:
        file=request.FILES['excel_file']
        obj=excel.objects.create(excel_file=file)
        name=str(file)
        l=name.split('.')
        if l[1]=='xlsx' or l[1]=='xls':


            #function to get random values of criteria
            def ra(value):
                if '±' in value:
                    p=value.split('±')
                    for i in range(len(p)):
                        p[i]=float(p[i])
                    x=round(random.uniform(p[0]-p[1],p[0]+p[1]), 2)
                    return x
                elif '+' in value:
                    p=value.split('+')
                    for i in range(len(p)):
                        p[i]=float(p[i])
                    x=round(random.uniform(p[0],p[0]+p[1]),2)
                    return x 
                elif '-' in value:
                    p=value.split('-')
                    for i in range(len(p)):
                        p[i]=float(p[i])
                    x=round(random.uniform(p[0]-p[1],p[0]),2)
                    return x 
                else:
                    return value





            doc=DocxTemplate('files/COA_format.docx')
            df = pd.read_excel(file)
            df1=pd.read_excel('files/COA_data.xlsx')
            #removing nan values and replacing with ' '
            s=df1.isnull().sum()
            df1=df1 .fillna(' ')
            materials=df1['MaterialName'].values
            para1=df1['para1'].values
            para2=df1['para2'].values
            para3=df1['para3'].values
            para4=df1['para4'].values
            para5=df1['para5'].values
            TestMethod1=df1['TestMethod1'].values
            TestMethod2=df1['TestMethod2'].values
            TestMethod3=df1['TestMethod3'].values
            TestMethod4=df1['TestMethod4'].values
            TestMethod5=df1['TestMethod5'].values
            Criteria1=df1['Criteria1'].values
            Criteria2=df1['Criteria2'].values
            Criteria3=df1['Criteria3'].values
            Criteria4=df1['Criteria4'].values
            Criteria5=df1['Criteria5'].values

            context={}

            for i in range(len(materials)):

    
                case={
            materials[i]:{
            'MaterialName':materials[i],
            'para1':para1[i],
            'para2':para2[i],
            'para3':para3[i],
            'para4':para4[i],
            'para5':para5[i],
            'TestMethod1':TestMethod1[i],
            'TestMethod2':TestMethod2[i],
            'TestMethod3':TestMethod3[i],
            'TestMethod4':TestMethod4[i],
            'TestMethod5':TestMethod5[i],
            'Criteria1':Criteria1[i],
            'Criteria2':Criteria2[i],
            'Criteria3':Criteria3[i],
            'Criteria4':Criteria4[i],
            'Criteria5':Criteria5[i],
            }
        }
    #a way to access the context2 with name of material in context
                context[materials[i]]=case
            context2={}
            c1=df['TestDate'].values
            c2 = df['COANo'].values
            c3 = df['MaterialName'].values
            c4 = df['BatchNo'].values
            c5 = df['Qty'].values
            c6 = df['ManuDate'].values
            c7 = df['ExpDate'].values
            c8 = df['InspectedBy'].values
            c9 = df['ApprovedBy'].values  
            for i in range(len(c1)):
    #converting numpy.datetime64 objects to date type
                test_date=pd.Timestamp(np.datetime64(c1[i])).to_pydatetime()
                test_date=str(test_date.date())
                test_date=datetime.datetime.strptime(test_date, "%Y-%m-%d").strftime("%d/%m/%Y").replace('/','-')
                exp_date=pd.Timestamp(np.datetime64(c7[i])).to_pydatetime()
                exp_date=str(exp_date.date())
                exp_date=datetime.datetime.strptime(exp_date, "%Y-%m-%d").strftime("%d/%m/%Y").replace('/','-')
                manu_date=pd.Timestamp(np.datetime64(c6[i])).to_pydatetime()
                manu_date=str(manu_date.date())
                manu_date=datetime.datetime.strptime(manu_date, "%Y-%m-%d").strftime("%d/%m/%Y").replace('/','-')
                context2={ 
                    'TestDate':test_date,
                    'COANo':c2[i],
                    'MaterialName':c3[i],
                    'BatchNo':c4[i],
                    'Qty':c5[i],
                    'ManuDate':manu_date,
                    'ExpDate':exp_date,
                    'InspectedBy':c8[i],
                    'ApprovedBy':c9[i],
                    'para1':context[c3[i]][c3[i]]['para1'],
                    'para2':context[c3[i]][c3[i]]['para2'],
                    'para3':context[c3[i]][c3[i]]['para3'],
                    'para4':context[c3[i]][c3[i]]['para4'],
                    'para5':context[c3[i]][c3[i]]['para5'],
                    'TestMethod1':context[c3[i]][c3[i]]['TestMethod1'],
                    'TestMethod2':context[c3[i]][c3[i]]['TestMethod2'],
                    'TestMethod3':context[c3[i]][c3[i]]['TestMethod3'],
                    'TestMethod4':context[c3[i]][c3[i]]['TestMethod4'],
                    'TestMethod5':context[c3[i]][c3[i]]['TestMethod5'],
                    'Criteria1':context[c3[i]][c3[i]]['Criteria1'],
                    'Criteria2':context[c3[i]][c3[i]]['Criteria2'],
                    'Criteria3':context[c3[i]][c3[i]]['Criteria3'],
                    'Criteria4':context[c3[i]][c3[i]]['Criteria4'],
                    'Criteria5':context[c3[i]][c3[i]]['Criteria5'],


                }
    #way of calling methods from format word file
                jinja_env = jinja2.Environment()
                jinja_env.filters['ra'] =ra
    


    #saving output files with desired names
                doc.render(context2,jinja_env)
                output_filename=f"output/{c2[i]}.docx"
                doc.save(output_filename)
    





        else:
            return HttpResponse("Please enter an excel file only")
        return render(request,'automater/home.html',{'file':obj})
    
    return render(request,'automater/home.html')
def download(request):
    l=glob.glob('output/*.docx')
    filenames = l


    zip_subdir = "Output-Word-Files"
    zip_filename = "%s.zip" % zip_subdir

 
    s = io.BytesIO()

    zf = zipfile.ZipFile(s, "w")

    for fpath in filenames:
        fdir, fname = os.path.split(fpath)
        zip_path = os.path.join(zip_subdir, fname)
        zf.write(fpath, zip_path)
    zf.close()
    resp = HttpResponse(s.getvalue(), content_type = "application/x-zip-compressed")
    resp['Content-Disposition'] = 'attachment; filename=%s' % zip_filename
    return resp