# File:"C:\Users\Aneesh\Documents\PTI\PSSE35\PV Analysis\PV Analysis Report\test1.py", generated on THU, NOV 20 2025  10:36, PSS(R)E Xplore release 35.06.01
import pssexcel
psspy.case(r"""savnw.sav""")
loads = [39.69,34.64,30.36,26.66,23.38,20.40,17.63,15,12.39,10.79,9.68,0]
for load in loads:
    psspy.case(r"""savnw.sav""")
    file_dfx= 'savnw_' + str(load) + '.dfx'
    file_pv = 'savnw_' + str(load) + '.pv'
    file_name_dfx = r"{}".format(file_dfx)
    file_name_pv = r"{}".format(file_pv)
    psspy.load_chng_6(205,r"""1""",[_i,_i,_i,_i,_i,_i,_i],[_f,load,_f,_f,_f,_f,_f,_f],"")
    psspy.fdns([0,0,0,0,1,1,99,0])
    psspy.dfax_2([1,1,0],r"""savnwSub.sub""",r"""savnwmon.mon""",r"""savnwcon.con""",file_name_dfx)
    psspy.pv_engine_6([0,1,0,1,0,0,0,2,0,0,1,1,4,1,0,0,0,0,0,0,1,0,0,0,0],[0.5,20.0,10.0,2500.0,0.9,100.0,0.0,0.0],[r"""SOURCE""",r"""SINK""",r"""SOURCE"""],
    file_name_dfx,"",r"""savnw.ecd""","",file_name_pv,"")
    pssexcel.pv(file_name_pv,['v','g','l'],colabel=['base case'],xlsfile='book.xlsx',sheet=str(load),overwritesheet=True,show=False)