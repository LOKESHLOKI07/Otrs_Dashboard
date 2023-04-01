import pandas as pd
import mysql.connector as sql
from plotly.offline import plot
import plotly.graph_objs as go


def get_ticket_duration_in_minutes(db_connection):
    data = pd.read_sql('SELECT ts.id as ticketstateID,ts.name AS ticketstateName,th.name,th.create_time,th.ticket_id  FROM ticket_history AS th \
                        left join ticket_state as ts ON th.state_id = ts.id \
                        WHERE history_type_id IN (27) \
                        AND th.ticket_id IN (1843719)', con=db_connection)

    fig = go.Figure(data=go.Bar(name='Plot1', x=data['ticket_id'], y=data['ticketstateName']))

    fig.update_layout(title_text='Plotly_Plot1',
                      xaxis_title='X_Axis',
                      yaxis_title='Y_Axis')

    plotly_plot_obj = plot({'data': fig}, output_type='div')

    return plotly_plot_obj


db_connection = sql.connect(host='otrs.futurenet.in', database='otrs5', user='readuser2', password='6FbUDa5VM')