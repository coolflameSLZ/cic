import os

from src.trans_value import Trans

if __name__ == '__main__':

    trans = Trans()

    for _, _, files in os.walk('../input'):
        for f in files:
            s_wb, t_wb = trans.genNewFile(f)

            trans.transing_工作记录单(s_wb, t_wb)
            trans.transing_水尺计算(s_wb, t_wb)
            trans.transing_压载水(s_wb, t_wb)

            s_wb.save()
            s_wb.close()

            t_wb.save()
            t_wb.close()

            print("{} 完成".format(f))
