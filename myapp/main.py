# Load data dari spreadsheet gdrive
import pandas as pd
import numpy as np
import datetime
from math import pi
from os.path import dirname, join
import calendar
from datetime import datetime

#Bokeh
from bokeh.io import push_notebook, show, output_notebook
from bokeh.core.properties import value
from bokeh.models.ranges import FactorRange
import math
from bokeh.palettes import Category20c, viridis, Reds256, cividis
from bokeh.layouts import row 
from bokeh.plotting import figure
from bokeh.models.widgets import DateRangeSlider
from bokeh.models import ColumnDataSource, HoverTool, Select, Slider, Div, CustomJS
from bokeh.transform import cumsum, factor_cmap
from bokeh.io import curdoc
from bokeh.layouts import column, row

import pandas as pd
sheet_id = "1Lb-8UYcZ5IIVyUS6dfPwWD6J7FQ0q-vAZgoK2jZXHK8"
sheet_url = "https://docs.google.com/spreadsheets/d/1wCRQ0ccLH-D9sFNvZh1t1Rk6085EVqU83vH5bj-I_cs/edit#gid=829398208"
url_1 = sheet_url.replace('/edit#gid=', '/export?format=csv&gid=')
data = pd.read_csv(url_1)

data.drop('NO', inplace=True, axis=1)

filename = join(dirname(__file__), "template/index.html")
desc = Div(text=open(filename).read(),
           render_as_text=False, width=1000)

cleaned_data = data.replace(r'^\s*$', np.nan, regex=True)
cleaned_data = cleaned_data.replace("-", np.nan)
cleaned_data = cleaned_data.replace("NULL", np.nan)
cleaned_data = cleaned_data.replace(("LME EXISTING","EXISTING","Existing"), np.nan)

cleaned_data['TANGGAL LME OK'] = cleaned_data['TANGGAL LME OK'].fillna(method='ffill')
cleaned_data['Drop Cable FO atas tanah / aerial 1 Core Single mode G.657'] = cleaned_data['Drop Cable FO atas tanah / aerial 1 Core Single mode G.657'].fillna(cleaned_data['Drop Cable FO atas tanah / aerial 1 Core Single mode G.657'].mode()[0])
cleaned_data['Splice on Connector Sumitomo'] = cleaned_data['Splice on Connector Sumitomo'].fillna(cleaned_data['Splice on Connector Sumitomo'].mode()[0])
cleaned_data['Precon KSO Indoor Trans 15 mtr dgn Roset'] = cleaned_data['Precon KSO Indoor Trans 15 mtr dgn Roset'].fillna(cleaned_data['Precon KSO Indoor Trans 15 mtr dgn Roset'].mode()[0])
cleaned_data['OTP FTTH 1 Port With Adaptor'] = cleaned_data['OTP FTTH 1 Port With Adaptor'].fillna(cleaned_data['OTP FTTH 1 Port With Adaptor'].mode()[0])
cleaned_data['S-Clamp-Springer'] = cleaned_data['S-Clamp-Springer'].fillna(cleaned_data['S-Clamp-Springer'].mode()[0])
cleaned_data['Breket / Clamp Hook'] = cleaned_data['Breket / Clamp Hook'].fillna(cleaned_data['Breket / Clamp Hook'].mode()[0])
cleaned_data['MITRA'] = cleaned_data['MITRA'].fillna(cleaned_data['MITRA'].mode()[0])
cleaned_data['STATUS LME'] = cleaned_data['STATUS LME'].fillna(cleaned_data['STATUS LME'].mode()[0])
s= cleaned_data[['KETERANGAN','ODP','BARCODE','TAGING LOKASI','SN ONT','GPON','PORT GPON','SN AP']]
cleaned_data[s.columns] = s.fillna('UNIDENTIFIED')

cleaned_data["TANGGAL LME OK"] = pd.to_datetime(cleaned_data["TANGGAL LME OK"], dayfirst=True)
cleaned_data["TANGGAL PERMINTAAN"] = pd.to_datetime(cleaned_data["TANGGAL PERMINTAAN"], dayfirst=True)
dates = []
for i in cleaned_data["TANGGAL PERMINTAAN"]:
    i = i.date()
    dates.append(i)  
    
cleaned_data["LAYANAN WMS"] = cleaned_data["LAYANAN WMS"].replace("WMS LITE", "LITE")
cleaned_data["LAYANAN WMS"] = cleaned_data["LAYANAN WMS"].replace("WMS Regular", "Reguler")
cleaned_data["LAYANAN WMS"] = cleaned_data["LAYANAN WMS"].replace("PDA WMS", "Reguler")
cleaned_data["LAYANAN WMS"] = cleaned_data["LAYANAN WMS"].replace("Regular", "Reguler")

# Text Total Order
# header_total_order = Div(text='<p style="text-align: center;">Total Order</p>')
text = len(cleaned_data['Status terakhir'])
text = """
            <div style="text-align: center;
                        box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);
                        transition: 0.3s;
                        margin-top: 20px;
                        margin-left: 20px;
                        background-color: #D61C4E;
                        padding: 25px 50px;">
                <h1 style="text-align:center;">{text}</h1>
                <div>
                    <p>Total order</p>
                </div>
            </div>
""".format(text=text)

div_order = Div(text=text)



x=cleaned_data['LAYANAN WMS'].value_counts()
data_layanan = pd.Series(x).reset_index(name='value').rename(columns={'index':'layanan'})

data_layanan['angle'] = data_layanan['value']/data_layanan['value'].sum() * 2*pi
data_layanan['color'] = cividis(len(x))
z=110*(data_layanan['value']/data_layanan['value'].sum())
data_layanan['value']=z

p = figure(title="Jenis Layanan", plot_height=300, toolbar_location=None ,
           tools="hover", tooltips="@layanan: @value{0.2f} %", x_range=(-.5, .5))
p.axis.visible = False
p.grid.grid_line_color = None
p.annular_wedge(x=0, y=1,  inner_radius=0.1, outer_radius=0.17, direction="anticlock",
                start_angle=cumsum('angle', include_zero=True), end_angle=cumsum('angle'),
        line_color="white", legend='layanan',fill_color='color', source=data_layanan)

#Plot 2
cleaned_data['year'] = cleaned_data['TANGGAL PERMINTAAN'].dt.year
cleaned_data['month'] = cleaned_data['TANGGAL PERMINTAAN'].dt.month
order_by_year = cleaned_data['year'].value_counts()
order_by_month = cleaned_data['month'].value_counts()
data_per_mount = pd.Series(order_by_month).reset_index(name='value').rename(columns={'index':'bulan'})
num_of_order = cleaned_data.groupby('TANGGAL PERMINTAAN').size().reset_index()
num_of_order.columns = ['Date', 'Jumlah']
num_of_order.set_index(num_of_order.Date, inplace=True)
num_of_order = num_of_order.resample('D').sum().fillna(0)
num_of_order.reset_index(inplace=True)
num_of_order2 = num_of_order
TOOLS = "pan, wheel_zoom, box_zoom, box_select,reset, save" # the tools you want to add to your graph
source_order = ColumnDataSource(data={'x':num_of_order['Date'], 'y':num_of_order['Jumlah']})
source_order2 = ColumnDataSource(data={'x':num_of_order['Date'], 'y':num_of_order['Jumlah']})
p_order = figure(title="Daily Order", x_axis_type='datetime',tools = TOOLS, toolbar_location=None, width=1000,height=350)
p_order.line(x='x', y='y', source=source_order, line_width=2, color="red") 
# p_order.xaxis.axis_label = 'Date'
# p_order.yaxis.axis_label = 'Jumlah Order'
p_order.xgrid.grid_line_color = None
p_order.legend.title_text_font_style = "bold"
p_order.legend.title_text_font_size = "20px"
p_order.border_fill_color = "white"
p_order.min_border_left = 10
p_order.min_border_right = 10
p_order.outline_line_width = 10
p_order.outline_line_alpha = 0.3
p_order.outline_line_color = "whitesmoke"
p_order.add_tools(HoverTool(tooltips=[('Date', '$x{%F}'),
                                ('Order', '@y')],
          formatters={'$x': 'datetime'}))
#Plot 3
cleaned_data["MITRA"] = cleaned_data["MITRA"].replace("MIB(AP)", "MIB")
cleaned_data["MITRA"] = cleaned_data["MITRA"].replace("PTJ(AP)", "PTJ")
#Handling outlier
cleaned_data.drop(cleaned_data[(cleaned_data['MITRA'] == 'Assurance')].index, inplace=True)
mitra = cleaned_data['MITRA'].value_counts()
data_mitra = pd.Series(mitra).reset_index(name='order').rename(columns={'index': 'mitra'})
data_mitra['angle'] = data_mitra['order']/data_mitra['order'].sum() * 2*pi
data_mitra['color'] = Category20c[len(mitra)]

p_mitra = figure(height=350, title="Mitra", toolbar_location=None,
           tools="hover", tooltips="@mitra: @order", x_range=(-0.6, 0.7))

p_mitra.wedge(x=0, y=1, radius=0.3,
        start_angle=cumsum('angle', include_zero=True), end_angle=cumsum('angle'),
        line_color="white", fill_color='color', legend_field='mitra', source=data_mitra)

p_mitra.axis.axis_label = None
p_mitra.axis.visible = False
p_mitra.min_border_top = 65
p_mitra.grid.grid_line_color = None

#Plot 4
cleaned_data['monthName'] = cleaned_data['month'].apply(lambda x: calendar.month_name[x])
num_of_order_by_month = cleaned_data.groupby([(cleaned_data.year), (cleaned_data.month),(cleaned_data.monthName)]).size().reset_index()
num_of_order_by_month.columns = ['Year','MonthNumber','Month', 'Count']
TOOLS = "pan, wheel_zoom, box_zoom, box_select,reset, save" 
source4 = ColumnDataSource(num_of_order_by_month)
monthList = source4.data['Month'].tolist()
p_order_per_month = figure(x_range=monthList, width=1200,height=300)
colorMonth = num_of_order_by_month['Month'].value_counts()
color_map = factor_cmap(field_name='Month',
                    palette=Category20c[len(colorMonth)], factors=monthList)

p_order_per_month.vbar(x='Month', top='Count', source=source4, width=0.80, color=color_map)

p_order_per_month.title.text ='Order PSB per Bulan'
p_order_per_month.xaxis.axis_label = 'Bulan'
p_order_per_month.yaxis.axis_label = 'Jumlah Order'

hover = HoverTool()
hover.tooltips = [
    ("Month", "@Month"),
    ('Orders','@Count')]

hover.mode = 'vline'

p_order_per_month.add_tools(hover)

best_order_data = num_of_order2[num_of_order2['Jumlah']==max(num_of_order2['Jumlah'])].index.values.astype(int)
best_order = num_of_order['Date'][best_order_data]
text = str(best_order)[4:14]

text = """
            <div style="text-align: center;
                        box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);
                        transition: 0.3s;
                        margin-left: 20px;
                        background-color: #CFD2CF;
                        padding: 25px 13px;">
                <h1 style="text-align:center;">{text}</h1>
                <div>
                    <p>Max Order</p>
                </div>
            </div>
""".format(text=text)

div_best_order = Div(text=text)

#Plot 5
df_order_2022 = cleaned_data[cleaned_data['monthName'] == 'July']
ct = pd.crosstab(df_order_2022['TANGGAL PERMINTAAN'], df_order_2022['MITRA'])
mitra_l = ct.columns.values
ct = ct.reset_index()
ct['TANGGAL PERMINTAAN'] = ct['TANGGAL PERMINTAAN'].astype(str)

source5 = ColumnDataSource(ct) # 
Date = source5.data['TANGGAL PERMINTAAN']
p_perform_mitra = figure(x_range=Date, title="Performansi Mitra",
           tools = TOOLS, width=900, height=400)

renderers = p_perform_mitra.vbar_stack(mitra_l, x='TANGGAL PERMINTAAN', source=source5, width=0.5, color=viridis(len(mitra_l)), 
             legend=[value(x) for x in mitra_l])

p_perform_mitra.xaxis.axis_label = 'Tanggal'
p_perform_mitra.yaxis.axis_label = 'Jumlah Penanganan Order'

p_perform_mitra.xgrid.grid_line_color = None

p_perform_mitra.y_range.start = 0
p_perform_mitra.y_range.end = 10 #to make room for the legend
p_perform_mitra.xgrid.grid_line_color = None
p_perform_mitra.axis.minor_tick_line_color = None
p_perform_mitra.outline_line_color = None

#add hover
hover = HoverTool()
hover.tooltips=[
    ('Mitra', '$name'), #$name provides data from legend
    ('Order', '@$name') #@$name gives the value corresponding to the legend
]
hover.formatters = {'Date': 'datetime'}
p_perform_mitra.add_tools(hover)
p_perform_mitra.xaxis.major_label_orientation = math.pi/2 

#CJs
date_slider = DateRangeSlider(
    title="Range Date", start=num_of_order['Date'][0], end=num_of_order['Date'][len(num_of_order)-1],
    value=(num_of_order['Date'][5], num_of_order['Date'][20]), step=1, width=900)

#num_of_order = num_of_order.to_json()
#num_of_order2 = num_of_order2.to_json()
callback = CustomJS(args=dict(source=source_order, ref_source=source_order2), code="""
    
    // print out array of date from, date to
    console.log(cb_obj.value); 
    
    // dates returned from slider are not at round intervals and include time;
    const date_from = Date.parse(new Date(cb_obj.value[0]).toDateString());
    const date_to = Date.parse(new Date(cb_obj.value[1]).toDateString());
    console.log(date_from, date_to)
    
    const data = source.data;
    const ref = ref_source.data;
    
    let new_ref = []
    ref["x"].forEach(elem => {
        elem = Date.parse(new Date(elem).toDateString());
        new_ref.push(elem);
        console.log(elem);
    })
    
    const from_pos = new_ref.indexOf(date_from);
    // add + 1 if you want inclusive end date
    const to_pos = new_ref.indexOf(date_to) + 1;
        
    // re-create the source data from "reference"
    data["y"] = ref["y"].slice(from_pos, to_pos);
    data["x"] = ref["x"].slice(from_pos, to_pos);
    
    source.change.emit();
    """)
    
date_slider.js_on_change('value', callback)

# Select
bulan= [
    'January',
    'February',
    'Maret',
    'April',
    'Mei',
    'Juni',
    'Juli',
    'Agustus',
    'November',
    'Desember'
]

def update_month(attrname, old, new):
    if month_select.value == 'January':
        newSource = cleaned_data[cleaned_data['monthName'] == 'January']
        newSourceCt = pd.crosstab(newSource['TANGGAL PERMINTAAN'], newSource['MITRA'])
    if month_select.value == 'February':
        newSource = cleaned_data[cleaned_data['monthName'] == 'February']
        newSourceCt = pd.crosstab(newSource['TANGGAL PERMINTAAN'], newSource['MITRA'])
    source5.data =  newSourceCt
    Date = source5.data['TANGGAL PERMINTAAN']
    
month_select = Select(value='January',
                          title='Pilih Bulan:',
                          width=300,
                          options=list(cleaned_data['monthName'].unique()))
year_select = Select(value='2021',
                    title='Pilih Tahun:',
                    width=300, 
                    options=['2021','2022'])
month_select.on_change('value', update_month)

copy_div = Div(text="""
           <div style="text-align: center;
                        padding: 10px 10px;">
                    <p>made with ❤️ by
                    <a href="https://www.linkedin.com/in/wida-sofiya">
                      Wida Sofiya
                    </a>&#127808
                  </p>
                <div>
                    <p></p>
                </div>
            </div>
""")

#Layout
app_title = desc
header_order = column(div_order, div_best_order)
widgets = row(p_order, header_order)

slide_date = column(widgets, date_slider)
series = row(p, p_mitra)
main_row = column(slide_date, series)
select = column(month_select, year_select)
mitra_row = row(select,p_perform_mitra)
#series = column(time_series1, time_series2)
layout = column(app_title, main_row, p_order_per_month, mitra_row, copy_div)


curdoc().add_root(column(layout, height=50))

curdoc().title = "Visualisasi PSB 2021-2022"
