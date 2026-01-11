import dearpygui.dearpygui as dpg
from math import sin, cos
import time,threading
import random

dpg.create_context()

time_ax = []
db_ay = []


def update_data():
    global sec_n
    sec_n = sec_n + 1
    db_val = random.randint(30,40)

    time_ax.append(sec_n)
    db_ay.append(db_val)

    dpg.set_value('series_tag', [time_ax, db_ay])
    dpg.set_item_label('series_tag', "nosie val:"+str(db_val)+":db")

def time_func():
    #update_series()
    update_data()
    timer_max_val = 1
    print("timer is occured")
    timer = threading.Timer(timer_max_val,time_func)
    timer.start()
with dpg.window(label="Noise DB", tag="win"):
    #dpg.add_button(label="Update Series", callback=update_series)
    # create plot
    timer_max_val = 1
    sec_n = 0
    timer = threading.Timer(timer_max_val,time_func)
    timer.start()
    with dpg.plot(label="Line Series", height=400, width=400):
        # optionally create legend
        dpg.add_plot_legend()
        # REQUIRED: create x and y axes
        dpg.add_plot_axis(dpg.mvXAxis, label="x")
        dpg.set_axis_limits(dpg.last_item(), 0, 2000)
        dpg.add_plot_axis(dpg.mvYAxis, label="y", tag="y_axis")
        dpg.set_axis_limits(dpg.last_item(), 0,100)

        # series belong to a y axis
        dpg.add_line_series(time_ax, db_ay, label="nosie db", parent="y_axis", tag="series_tag")

dpg.create_viewport(title='Environment Detection', width=800, height=600)
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.start_dearpygui()
dpg.destroy_context()