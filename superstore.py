import dash
from dash import html
from dash import dcc
import dash_bootstrap_components as dbc

import  plotly.express as px
import pandas as pd

df = pd.read_excel('Super Store.xlsx', index_col='Order Date', sheet_name='Orders')
graph4DF = df[['Year (OrderDate)', 'Profit', 'Market']].groupby(['Market','Year (OrderDate)']).sum()
graph4DF_years = [t[1] for t in graph4DF.index]
graph4DF_market = [t[0] for t in graph4DF.index]


graph3DF = df[['Region', 'Sales', 'Target Sales']].groupby('Region').sum()

graph2DF = df[['Customer ID', 'Customer Name', 'Target Sales']].groupby(['Customer ID', 'Customer Name']).sum().sort_values('Target Sales', ascending=False).head(10).sort_values('Target Sales')
graph2DF_names = [t[1] for t in graph2DF.index]

graph1DF = df[['Profit']]
graph1DF['Month'] = graph1DF.index.strftime('%b')
graph1DF =graph1DF[graph1DF.index.year == graph1DF.index.year.max() ]
graph1DF = graph1DF.groupby('Month').sum()
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
graph1DF = graph1DF.reindex(months, axis=0)

no_of_customers = df['Customer ID'].value_counts().count()

df2 = df[['Year (OrderDate)', 'Profit']].groupby('Year (OrderDate)').sum()
inc_percentage = int((df2.iloc[-1, 0] - df2.iloc[-2, 0])*100/df2.iloc[-2, 0])

df3 = df[df['Year (OrderDate)'] == df['Year (OrderDate)'].max()][['Country', 'Profit']].groupby('Country').sum()
avg = df3['Profit'].mean()
Poor_countries = df3[df3['Profit'] < avg]['Profit'].count()


app = dash.Dash(external_stylesheets=[dbc.themes.BOOTSTRAP])
app.layout=html.Div([
        html.H1('Superstore', style={'margin-bottom':'35px', 'text-align':'center', 'color':'grey', 'text-shadow':'1px 1px black'}),
        
        html.Div([
                html.Div([
                        html.H3(str(inc_percentage) + '%', style={'color':'#636EFA', 'font-size':'45px', 'font-family': 'fantasy'}),
                        html.P('Percentage increase in Profit', style={'font-size':'12px', 'color':'#636EFA'})
                    ], className='col-md-4', style={'text-align':'center'}),
                html.Div([
                        html.H3(str(Poor_countries), style={'color':'#636EFA', 'font-size':'45px', 'font-family': 'fantasy'}),
                        html.P('No. of countries with total profit less than the average', style={'font-size':'12px', 'color':'#636EFA'})
                        
                    ], className='col-md-4', style={'text-align':'center'}),
                html.Div([
                        html.H3(str(no_of_customers), style={'color':'#636EFA', 'font-size':'45px', 'font-family': 'fantasy'}),
                        html.P('Total No. Of Customers', style={'font-size':'12px', 'color':'#636EFA'})
                        
                    ], className='col-md-4', style={'text-align':'center'})
            ], className='row'),
        html.Div([
                html.Div([
                        dcc.Graph(id = 'plot1',
                                  figure=px.line(graph1DF, y='Profit',title="Change in Profit accross last year"))
                    ], className='col-md-8'),
                html.Div([
                        dcc.Graph(id = 'plot2',
                                  figure=px.bar(graph2DF, x='Target Sales', y=graph2DF_names, orientation='h', title="Top Customers",
                                                labels={"y":"Customers"}))
                    ], className='col-md-4')
            ], className='row'),
        html.Div([
                html.Div([
                        dcc.Graph(id = 'plot3',
                                  figure=px.area(graph3DF, y=['Sales', 'Target Sales'], title="Region Comparison according to sales"))
                    ], className='col-md-8'),
                html.Div([
                        dcc.Graph(id = 'plot4',
                                  figure=px.line(graph4DF, x=graph4DF_years, y='Profit', color=graph4DF_market, title="Difference in market share",
                                                 labels={"x":"years", "y":"Profit", "color":"Market"}))
                    ], className='col-md-4')
            ], className='row')
    ], className='container')

app.run_server()