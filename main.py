import datetime
import sys
import psutil
import wmi
from PySide6 import QtCharts
from PySide6.QtCore import QDateTime, QTimer
from PySide6.QtGui import Qt
from PySide6.QtWidgets import QWidget, QApplication

from PC_Monitor import Ui_PC_Monitor

import win32com.client

def get_total_thread_count_pywin32():
    """使用 pywin32 获取系统总线程数"""
    wmi = win32com.client.GetObject("winmgmts:")
    processes = wmi.InstancesOf("Win32_Process")
    return sum(int(p.Properties_("ThreadCount").Value) for p in processes)


def get_total_handle_count():
    """获取系统中所有进程的句柄总数"""
    total_handles = 0
    for proc in psutil.process_iter(['pid', 'name', 'num_handles']):
        try:
            total_handles += proc.info['num_handles']
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return total_handles




class MyWindow(QWidget):
    def __init__(self,parent=None):
        super().__init__(parent)
        self.c = wmi.WMI()

        self.ui = Ui_PC_Monitor()  # 创建UI对象
        self.ui.setupUi(self) # 构造UI界面
        self.create_chart()
        self.set_timer()


        self.setWindowTitle("电脑系统监控")


    def create_chart(self):
        self.chart = QtCharts.QChart()
        self.series = QtCharts.QLineSeries()
        self.chart.addSeries(self.series)
        self.chart.legend().hide()

        #设置坐标轴显示范围
        self.data_axisX = QtCharts.QDateTimeAxis()
        self.value_axisY = QtCharts.QValueAxis()

        self.limitminute = 1 #设置显示多少分钟内活动
        self.maxspeed = 100 # 预设y轴最大值
        #Returns a QDateTime object containing a datetime a seconds later than the datetime of this object (or earlier if s is negative).
        self.data_axisX.setMin(QDateTime.currentDateTime().addSecs(-self.limitminute*60))
        self.data_axisX.setMax(QDateTime.currentDateTime().addSecs(0))
        self.value_axisY.setMin(0)
        self.value_axisY.setMax(self.maxspeed)
        self.data_axisX.setFormat("hh:mm:ss")
        #把坐标轴添加到 chart 中
        self.chart.addAxis(self.data_axisX,Qt.AlignmentFlag.AlignBottom)
        self.chart.addAxis(self.value_axisY,Qt.AlignmentFlag.AlignLeft)

        # 把曲线关联到坐标轴
        self.series.attachAxis(self.data_axisX)
        self.series.attachAxis(self.value_axisY)
        self.ui.graphicsView.setChart(self.chart)

        self.ui.cpu_L2_cache_value.setText(str(self.c.Win32_Processor()[0].L2CacheSize / 1024)+' MB')
        self.ui.cpu_L3_cache_value.setText(str(self.c.Win32_Processor()[0].L3CacheSize / 1024)+' MB')
        # print("系统名称: "+self.c.Win32_OperatingSystem()[0].Caption)

        cpufreq = psutil.cpu_freq()
        self.ui.cpu_kernel_value.setText(str(psutil.cpu_count(logical=False)))
        self.ui.cpu_logicprocessor_value.setText(str(psutil.cpu_count(logical=True)))
        self.ui.cpu_basespeed_value.setText(str('%.2f'% (cpufreq.current/1000))+' GHz')
        self.ui.cpu_name.setText(str(self.c.Win32_Processor()[0].Name))


    def set_timer(self):
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.cpuLoad)
        self.timer.start(1000)  # 每隔200毫秒出一个点

    def cpuLoad(self):
        current_time = QDateTime.currentDateTime()
        self.data_axisX.setMin(current_time.addSecs(-self.limitminute*60))
        self.data_axisX.setMax(current_time.addSecs(0))
        cpuload = psutil.cpu_percent()
        self.series.append(current_time.toMSecsSinceEpoch(),cpuload)
        if self.series.at(0):
            # Returns the datetime as the number of milliseconds that have passed since 1970-01-01T00:00:00.000,Coordinated Universal Time(UTC)
            if self.series.at(0).x()<current_time.addSecs(-self.limitminute*60).toMSecsSinceEpoch():
                self.series.remove(0)
        self.ui.cpu_percent_value.setText(str(cpuload)+'%')
        self.ui.cpu_runningtime_value.setText(str(datetime.datetime.fromtimestamp(psutil.boot_time()).strftime("%d:%H:%M:%S")))
        self.ui.cpu_process_value.setText(str(len(psutil.pids())))

        self.ui.cpu_thread_value.setText(str(get_total_thread_count_pywin32()))

        self.ui.cpu_handle_value.setText(str(get_total_handle_count()))






if  __name__ == '__main__':
    app = QApplication(sys.argv)

    window = MyWindow()
    window.show()
    sys.exit(app.exec())

 