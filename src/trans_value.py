import shutil

import xlwings as xw
import datetime as dt

class Trans:

    def genNewFile(self, source_file_name):
        target_file_name = '[B]_{}'.format(source_file_name)
        target_file_fullpath = '../output/{}'.format(target_file_name)

        # 拷贝
        shutil.copy('../resources/空白水尺模板.xlsx', target_file_fullpath)

        # 打开两个文件并返回
        s_wb = xw.Book('../input/{}'.format(source_file_name))
        t_wb = xw.Book(target_file_fullpath)

        return s_wb, t_wb


    def transing_工作记录单(self, s_wb, t_wb):


        s_sheet_工作记录 = s_wb.sheets['工作记录']
        s_sheet_货物信息 = s_wb.sheets['货物信息']
        s_sheet_记1 = s_wb.sheets['记(1)']
        t_sheet_工作记录单 = t_wb.sheets['工作记录单']

        #%%

        # s_sheet_工作记录
        v_船舶名称 = s_sheet_工作记录.range('C18').value
        v_船次 = s_sheet_工作记录.range('C19').value
        v_靠泊时间 = s_sheet_工作记录.range('C20').options(dates=dt.datetime).value
        v_注册国家地区 = s_sheet_工作记录.range('C21').value
        v_首次_工作人员 = s_sheet_工作记录.range('C5').value
        v_首次_工作日期 = s_sheet_工作记录.range('C6').value
        v_首次_工作时间 = s_sheet_工作记录.range('C7').value
        v_首次_工作地点 = s_sheet_工作记录.range('C8').value
        v_末次_工作人员 = s_sheet_工作记录.range('D5').value
        v_末次_工作日期 = s_sheet_工作记录.range('D6').value
        v_末次_工作时间 = s_sheet_工作记录.range('D7').value
        v_末次_工作地点 = s_sheet_工作记录.range('D8').value
        v_异常情况 = s_sheet_工作记录.range('C13').value
        v_备注 = s_sheet_工作记录.range('C14').value
        v_船龄 = s_sheet_记1.range('J7').value
        v_船龄 = str(v_船龄).replace('年', '')
        v_装卸始 = '{} {}'.format(v_首次_工作日期.strftime('%Y/%m/%d'), v_首次_工作时间.split('-')[1])
        #%%
        # s_sheet_货物信息
        v_货物卸毕时间 = s_sheet_货物信息.range('G7').value  #注册国家地区

        t_sheet_工作记录单.range('C7').value = v_船舶名称
        t_sheet_工作记录单.range('C8').value = v_船次
        t_sheet_工作记录单.range('F8').value = v_注册国家地区
        t_sheet_工作记录单.range('N8').value = v_船龄
        t_sheet_工作记录单.range('D10').value = (v_靠泊时间 - dt.timedelta(days=1))
        t_sheet_工作记录单.range('F10').value = (v_靠泊时间)
        t_sheet_工作记录单.range('L10').value = v_装卸始
        t_sheet_工作记录单.range('R10').value = v_货物卸毕时间
        t_sheet_工作记录单.range('D11').value = v_首次_工作人员
        t_sheet_工作记录单.range('D12').value = v_首次_工作日期
        t_sheet_工作记录单.range('E12').value = v_首次_工作时间
        t_sheet_工作记录单.range('D13').value = v_首次_工作地点
        t_sheet_工作记录单.range('L11').value = v_末次_工作人员
        t_sheet_工作记录单.range('L12').value = v_末次_工作日期
        t_sheet_工作记录单.range('N12').value = v_末次_工作时间
        t_sheet_工作记录单.range('L13').value = v_末次_工作地点
        t_sheet_工作记录单.range('B26').value = '备注：{}  异常情况：{}'.format(v_备注, v_异常情况)

        print("工作记录单 完成")

    def transing_水尺计算(self, s_wb, t_wb):

        s_sheet_水尺计算 = s_wb.sheets['水尺计算']
        s_sheet_船用物料 = s_wb.sheets['船用物料']
        t_sheet_水尺计算初次 = t_wb.sheets['水尺计算初次']
        t_sheet_水尺计算末次 = t_wb.sheets['水尺计算末次']

        # s_sheet_水尺计算
        v_首次_艏左舷 = s_sheet_水尺计算.range('D12').value
        v_首次_艏右舷 = s_sheet_水尺计算.range('D13').value
        v_首次_舯左舷 = s_sheet_水尺计算.range('D15').value
        v_首次_舯右舷 = s_sheet_水尺计算.range('D16').value
        v_首次_艉左舷 = s_sheet_水尺计算.range('D18').value
        v_首次_艉右舷 = s_sheet_水尺计算.range('D19').value
        v_首次_LBP = s_sheet_水尺计算.range('D5').value
        v_首次_艏垂线距 = s_sheet_水尺计算.range('D23').value
        v_首次_舯垂线距 = s_sheet_水尺计算.range('D24').value
        v_首次_艉垂线距 = s_sheet_水尺计算.range('D25').value
        v_首次_TPC = s_sheet_水尺计算.range('D42').value
        v_首次_LCF = s_sheet_水尺计算.range('D43').value
        v_首次_MTC1 = s_sheet_水尺计算.range('D44').value
        v_首次_MTC2 = s_sheet_水尺计算.range('D45').value
        v_首次_港水密度 = s_sheet_水尺计算.range('D39').value
        v_首次_空船重量 = s_sheet_水尺计算.range('D7').value
        v_首次_预报船舶常数 = s_sheet_水尺计算.range('D8').value
        v_首次_查表水尺 = s_sheet_水尺计算.range('D40').value
        v_首次_查表排水量 = s_sheet_水尺计算.range('D41').value

        #====================================【末次】==================================

        v_末次_艏左舷 = s_sheet_水尺计算.range('E12').value
        v_末次_艏右舷 = s_sheet_水尺计算.range('E13').value
        v_末次_舯左舷 = s_sheet_水尺计算.range('E15').value
        v_末次_舯右舷 = s_sheet_水尺计算.range('E16').value
        v_末次_艉左舷 = s_sheet_水尺计算.range('E18').value
        v_末次_艉右舷 = s_sheet_水尺计算.range('E19').value
        v_末次_LBP = s_sheet_水尺计算.range('D5').value
        v_末次_艏垂线距 = s_sheet_水尺计算.range('E23').value
        v_末次_舯垂线距 = s_sheet_水尺计算.range('E24').value
        v_末次_艉垂线距 = s_sheet_水尺计算.range('E25').value
        v_末次_TPC = s_sheet_水尺计算.range('E42').value
        v_末次_LCF = s_sheet_水尺计算.range('E43').value
        v_末次_MTC1 = s_sheet_水尺计算.range('E44').value
        v_末次_MTC2 = s_sheet_水尺计算.range('E45').value
        v_末次_港水密度 = s_sheet_水尺计算.range('E39').value
        v_末次_空船重量 = s_sheet_水尺计算.range('D7').value
        v_末次_预报船舶常数 = s_sheet_水尺计算.range('D8').value
        v_末次_查表水尺 = s_sheet_水尺计算.range('E40').value
        v_末次_查表排水量 = s_sheet_水尺计算.range('E41').value

        # s_sheet_船用物料
        v_首次_重油 = s_sheet_船用物料.range('F7').value
        v_首次_清油 = s_sheet_船用物料.range('F8').value
        v_首次_滑油 = s_sheet_船用物料.range('F9').value
        v_末次_重油 = s_sheet_船用物料.range('H7').value
        v_末次_清油 = s_sheet_船用物料.range('H8').value
        v_末次_滑油 = s_sheet_船用物料.range('H9').value

        #==================================【首次】=====================================

        t_sheet_水尺计算初次.range('D9').value = v_首次_艏左舷
        t_sheet_水尺计算初次.range('L9').value = v_首次_艏右舷
        t_sheet_水尺计算初次.range('D10').value = v_首次_舯左舷
        t_sheet_水尺计算初次.range('L10').value = v_首次_舯右舷
        t_sheet_水尺计算初次.range('D11').value = v_首次_艉左舷
        t_sheet_水尺计算初次.range('L11').value = v_首次_艉右舷
        t_sheet_水尺计算初次.range('C14').value = v_首次_LBP
        t_sheet_水尺计算初次.range('C15').value = v_首次_艏垂线距
        t_sheet_水尺计算初次.range('C16').value = v_首次_舯垂线距
        t_sheet_水尺计算初次.range('C17').value = v_首次_艉垂线距
        t_sheet_水尺计算初次.range('B21').value = v_首次_TPC
        t_sheet_水尺计算初次.range('G21').value = v_首次_LCF
        t_sheet_水尺计算初次.range('I21').value = v_首次_MTC1
        t_sheet_水尺计算初次.range('O21').value = v_首次_MTC2
        t_sheet_水尺计算初次.range('L25').value = v_首次_港水密度
        t_sheet_水尺计算初次.range('D34').value = v_首次_空船重量
        t_sheet_水尺计算初次.range('AC18').value = v_首次_查表水尺
        t_sheet_水尺计算初次.range('AE18').value = v_首次_查表排水量
        t_sheet_水尺计算初次.range('AD31').value = v_首次_预报船舶常数
        t_sheet_水尺计算初次.range('AC27').value = v_首次_重油
        t_sheet_水尺计算初次.range('AC28').value = v_首次_清油
        t_sheet_水尺计算初次.range('AC29').value = v_首次_滑油

        #==================================【末次】=====================================

        t_sheet_水尺计算末次.range('D9').value = v_末次_艏左舷
        t_sheet_水尺计算末次.range('L9').value = v_末次_艏右舷
        t_sheet_水尺计算末次.range('D10').value = v_末次_舯左舷
        t_sheet_水尺计算末次.range('L10').value = v_末次_舯右舷
        t_sheet_水尺计算末次.range('D11').value = v_末次_艉左舷
        t_sheet_水尺计算末次.range('L11').value = v_末次_艉右舷
        t_sheet_水尺计算末次.range('C14').value = v_末次_LBP
        t_sheet_水尺计算末次.range('C15').value = v_末次_艏垂线距
        t_sheet_水尺计算末次.range('C16').value = v_末次_舯垂线距
        t_sheet_水尺计算末次.range('C17').value = v_末次_艉垂线距
        t_sheet_水尺计算末次.range('B21').value = v_末次_TPC
        t_sheet_水尺计算末次.range('G21').value = v_末次_LCF
        t_sheet_水尺计算末次.range('I21').value = v_末次_MTC1
        t_sheet_水尺计算末次.range('O21').value = v_末次_MTC2
        t_sheet_水尺计算末次.range('L25').value = v_末次_港水密度
        t_sheet_水尺计算末次.range('D34').value = v_末次_空船重量
        t_sheet_水尺计算末次.range('AC18').value = v_末次_查表水尺
        t_sheet_水尺计算末次.range('AE18').value = v_末次_查表排水量
        t_sheet_水尺计算末次.range('AD31').value = v_末次_预报船舶常数
        t_sheet_水尺计算末次.range('AC27').value = v_末次_重油
        t_sheet_水尺计算末次.range('AC28').value = v_末次_清油
        t_sheet_水尺计算末次.range('AC29').value = v_末次_滑油

        print("水尺计算 完成")

    def transing_压载水(self, s_wb, t_wb):

        s_sheet_船用物料 = s_wb.sheets['船用物料']
        t_sheet_压载水 = t_wb.sheets['压载水']

        v_首次_FWTP = s_sheet_船用物料.range('F13').value
        v_首次_FWTS = s_sheet_船用物料.range('F14').value
        v_末次_FWTP = s_sheet_船用物料.range('H13').value
        v_末次_FWTS = s_sheet_船用物料.range('H14').value

        # s_sheet_水尺计算
        v_船舱型号 = s_sheet_船用物料.range('C20:C44').value
        v_船舱型号 = list(filter(None, v_船舱型号))
        print(v_船舱型号)
        v_首次_管高 = s_sheet_船用物料.range('D20:D44').value
        v_首次_测深 = s_sheet_船用物料.range('E20:E44').value
        v_首次_体积 = s_sheet_船用物料.range('F20:F44').value
        v_末次_测深 = s_sheet_船用物料.range('G20:G44').value
        v_末次_体积 = s_sheet_船用物料.range('H20:H44').value


        t_sheet_压载水.range('G37').value = v_首次_FWTP
        t_sheet_压载水.range('G38').value = v_首次_FWTS
        t_sheet_压载水.range('M37').value = v_末次_FWTP
        t_sheet_压载水.range('M38').value = v_末次_FWTS

        t_sheet_压载水.range('A11').options(transpose=True).value = v_船舱型号
        t_sheet_压载水.range('B11').options(transpose=True).value = v_首次_管高
        t_sheet_压载水.range('C11').options(transpose=True).value = v_首次_管高
        t_sheet_压载水.range('D11').options(transpose=True).value = v_首次_测深
        t_sheet_压载水.range('E11').options(transpose=True).value = v_首次_体积
        t_sheet_压载水.range('F11').options(transpose=True).value = [1.025 for i in range(len(v_船舱型号))]
        t_sheet_压载水.range('I11').options(transpose=True).value = v_末次_测深
        t_sheet_压载水.range('J11').options(transpose=True).value = v_末次_测深
        t_sheet_压载水.range('K11').options(transpose=True).value = v_末次_体积
        t_sheet_压载水.range('L11').options(transpose=True).value = [1.025 for i in range(len(v_船舱型号))]


        print("压载水 完成")
