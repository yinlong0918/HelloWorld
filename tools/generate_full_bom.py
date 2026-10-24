from pathlib import Path
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]

# Prices are RMB, estimated ex-VAT procurement prices as of 2026-10 for planning.
# They are not supplier quotations. DNP items have an extension cost of zero.
rows = [
 ["电源输入","J1","2P接线端子",1,"Degson/同等","DG128-5.08-02P","2Pin, 5.08mm, 300V/10A","THT 5.08mm","装配",1.20,0.85,"12V/GND"],
 ["电源输入","F1","自恢复保险丝",1,"Bourns","MF-MSMF050-2","0.5A hold, 1A trip, 15V","1812","装配",1.10,0.78,"车载版建议核对最大工作电压"],
 ["电源输入","D1","反接串联肖特基",1,"Vishay","SS110-E3/61T","1A, 100V Schottky","SMA","装配",0.55,0.38,"低功耗方案；更低压降可改理想二极管"],
 ["电源输入","TVS1","电源TVS",1,"Littelfuse","SMBJ24A","24V standoff, 600W","SMB","装配",1.80,1.25,"需结合ISO 7637-2脉冲重新核算"],
 ["电源输入","C1","输入陶瓷电容",1,"Murata","GRM32ER71H106KA12L","10uF, 50V, X7R, ±10%","1210","装配",1.80,1.20,"靠近U1 VIN"],
 ["电源输入","C2","输入高频去耦",1,"Murata","GRM21BR71H104KA01L","100nF, 50V, X7R","0805","装配",0.18,0.10,""],
 ["5V降压","U1","宽压降压转换器",1,"Texas Instruments","LMR16006YQDDCRQ1","4-60V input, 0.6A, AEC-Q100","SOT-23-6","装配",6.80,5.20,"正式投板须按TI参考设计复核补偿与热设计"],
 ["5V降压","L1","功率电感",1,"Sunlord","SWPA5040S470MT","47uH, Isat≥1.1A, DCR≤0.45Ω","5.0x5.0mm","装配",1.80,1.15,"可替换TDK/伍尔特同规格"],
 ["5V降压","D2","续流肖特基",1,"Vishay","SS36-E3/57T","3A, 60V Schottky","SMA","装配",0.75,0.48,"阴极接SW"],
 ["5V降压","C3","5V输出电容",1,"Murata","GRM32ER71A476KE15L","47uF, 10V, X7R","1210","装配",2.20,1.55,"核对DC bias后有效容量"],
 ["5V降压","C4","5V高频去耦",1,"Murata","GRM188R71A104KA01D","100nF, 10V, X7R","0603","装配",0.08,0.04,""],
 ["5V降压","C20","BOOT电容",1,"Murata","GRM188R71H104KA93D","100nF, 50V, X7R","0603","装配",0.10,0.05,"BOOT-SW"],
 ["5V降压","R1","反馈上拉电阻",1,"Yageo","RC0603FR-0756KL","56kΩ, 1%, 0.1W","0603","装配",0.05,0.02,"输出约5.0V，最终按U1基准核对"],
 ["5V降压","R2","反馈下拉电阻",1,"Yageo","RC0603FR-0710KL","10kΩ, 1%, 0.1W","0603","装配",0.05,0.02,""],
 ["3V3稳压","U2","3.3V LDO",1,"Texas Instruments","TLV70033QDDCRQ1","3.3V, 200mA, AEC-Q100","SOT-23-5","装配",2.20,1.55,""],
 ["3V3稳压","C5,C6","LDO输入/输出电容",2,"Murata","GRM21BR71A106KE51L","10uF, 10V, X7R","0805","装配",0.42,0.25,"每颗稳压器引脚附近放置"],
 ["3V3稳压","C7,C8,C9,C10,C11,C12,C14,C17,C18","数字/收发器去耦",9,"Murata","GRM188R71C104KA01D","100nF, 16V, X7R","0603","装配",0.07,0.035,"每个VDD/VCC/VIO引脚就近一颗"],
 ["MCU","U3","微控制器",1,"STMicroelectronics","STM32F103C8T6","Cortex-M3, 64KB Flash, bxCAN, LQFP48","LQFP-48","装配",12.50,9.50,"非车规；若正式装车需改AEC-Q100 MCU"],
 ["MCU","Y1","主晶振",1,"YXC","X322508MSB4SI","8MHz, 10ppm, CL=12pF","3225-4P","装配",0.95,0.62,"按实际CL调整负载电容"],
 ["MCU","C15,C16","晶振负载电容",2,"Murata","GRM1885C1H180JA01D","18pF, 50V, C0G, ±5%","0603","装配",0.08,0.04,"需结合PCB寄生与Y1 CL复核"],
 ["MCU","C13","MCU低频储能",1,"Murata","GRM188R71C105KA12D","1uF, 16V, X7R","0603","装配",0.12,0.07,""],
 ["MCU","R7","NRST上拉",1,"Yageo","RC0603FR-0710KL","10kΩ, 1%","0603","装配",0.05,0.02,""],
 ["MCU","R8","BOOT0下拉",1,"Yageo","RC0603FR-07100KL","100kΩ, 1%","0603","装配",0.05,0.02,""],
 ["MCU","SW1","复位按键",1,"ALPS Alpine/同等","SKRPACE010","轻触按键, 常开","SMD 4.2x3.2mm","装配",0.65,0.42,""],
 ["MCU","SW2","学习按键",1,"ALPS Alpine/同等","SKRPACE010","轻触按键, 常开","SMD 4.2x3.2mm","装配",0.65,0.42,""],
 ["MCU","D3","状态LED",1,"Everlight","19-217/GHC-YR1S2/3T","Green, 20mA","0603","装配",0.12,0.07,""],
 ["MCU","R9","LED限流",1,"Yageo","RC0603FR-071KL","1kΩ, 1%","0603","装配",0.05,0.02,""],
 ["MCU","R10","学习键上拉/下拉",1,"Yageo","RC0603FR-0710KL","10kΩ, 1%","0603","装配",0.05,0.02,"按固件有效电平装配"],
 ["调试","J4","SWD调试连接器",1,"Ckmtw/同等","1.27-5P-SMD","5Pin, 1.27mm","SMD","装配",0.80,0.48,"3V3/SWDIO/SWCLK/NRST/GND"],
 ["调试","TP1-TP5","测试点",5,"Keystone/国产","5015-equivalent","1.0mm test pad / loop","SMD/THT","装配",0.22,0.12,"3V3,GND,RF_RX,CAN_RX,CAN_TX"],
 ["RF接口","J2","433MHz接收模块连接器",1,"Ckmtw/同等","A2541WV-4P","1x4, 2.54mm","THT","装配",0.45,0.25,"5V/GND/UART_TX/NC；以实物引脚为准"],
 ["RF接口","RF1","433MHz UART接收模块",1,"供应商待锁定","433.92MHz UART RX","超外差/OOK，UART 9600，输出地址+键值","模块","装配",18.00,12.00,"必须先做样品协议、电平和尺寸验证"],
 ["RF接口","ANT1","433MHz天线",1,"国产/同等","433MHz spring antenna","433.92MHz, 50Ω或17.3cm线天线","THT","装配",1.50,0.85,"外壳为金属时改外置天线"],
 ["RF接口","ESD1","UART ESD保护",1,"Nexperia","PESD5V0S1BA","5V, low capacitance","SOD-323","装配",0.38,0.24,"靠近J2"],
 ["RF接口","R3","UART串联电阻",1,"Yageo","RC0603FR-071KL","1kΩ, 1%","0603","装配",0.05,0.02,""],
 ["RF接口","R4","UART下拉",1,"Yageo","RC0603FR-07100KL","100kΩ, 1%","0603","装配",0.05,0.02,""],
 ["RF接口","R5","5V转3V3分压上臂",1,"Yageo","RC0603FR-0710KL","10kΩ, 1%","0603","条件装配",0.05,0.02,"仅模块输出5V时装配"],
 ["RF接口","R6","5V转3V3分压下臂",1,"Yageo","RC0603FR-0720KL","20kΩ, 1%","0603","条件装配",0.05,0.02,"仅模块输出5V时装配"],
 ["RF接口","JP1","3V3 UART旁路焊桥",1,"PCB feature","SolderJumper-2","2Pad solder jumper","PCB","DNP",0.00,0.00,"仅确认模块输出3.3V时短接"],
 ["CAN","U4","高速CAN收发器",1,"Texas Instruments","TCAN1042HGVDRQ1","5V VCC, 3.3V VIO, ±70V bus fault, AEC-Q100","SOIC-8","装配",6.20,4.60,""],
 ["CAN","R11","STB下拉",1,"Yageo","RC0603FR-0710KL","10kΩ, 1%","0603","装配",0.05,0.02,"默认Normal模式"],
 ["CAN","L2","CAN共模电感",1,"TDK","ACT45B-510-2P-TL003","51uH, CAN/CAN-FD, AEC-Q200","1812-class","装配",4.20,3.10,"首板同时预留0Ω旁路"],
 ["CAN","JP2","CMC旁路",2,"Yageo","RC0805JR-070RL","0Ω jumper","0805","DNP",0.04,0.015,"仅不装L2时装配两颗"],
 ["CAN","D4","CAN双线TVS",1,"Littelfuse","SM24CANB-02HTG","24V, dual CAN line protection","SOT-23","装配",3.20,2.20,"靠近J3"],
 ["CAN","J3","CAN连接器",1,"Degson/同等","DG128-5.08-04P","4Pin, 5.08mm","THT","装配",2.20,1.45,"CANH/CANL/GND/SHIELD"],
 ["CAN","R12,R13","分裂终端电阻",2,"Yageo","RC0805FR-0760R4L","60.4Ω, 1%, 0.125W","0805","DNP",0.08,0.04,"仅总线物理末端装配"],
 ["CAN","C19","分裂终端电容",1,"Murata","GRM21BR71H472KA01L","4.7nF, 50V, X7R","0805","DNP",0.12,0.07,"仅总线物理末端装配"],
 ["CAN","JP3","终端使能焊桥",1,"PCB feature","SolderJumper-2","2Pad solder jumper","PCB","DNP",0.00,0.00,"可由0Ω电阻替代"],
 ["遥控器","RC1","15键433MHz遥控器",1,"供应商待锁定","EV1527-compatible 15-key","433.92MHz OOK, unique address","成品","装配",22.00,15.00,"RF433LED仅作搜索词，量产必须锁型号"],
 ["结构","H1-H4","M3安装孔/螺丝",4,"国产/同等","M3x6 + spacer","M3铜柱/螺丝","机械件","装配",0.35,0.22,""],
 ["结构","PCB1","双层PCB",1,"JLCPCB/同等","FR-4 1.6mm","2 layer, 1oz, HASL-LF, 约117x60mm","PCB","装配",8.00,4.50,"缩板后成本可下降"],
 ["结构","CASE1","塑料外壳",1,"Takachi/国产","约120x70x30mm","ABS, 阻燃建议UL94-V0","机械件","装配",18.00,11.00,"不开模通用外壳"],
 ["制造","SMT","SMT贴片及手焊",1,"EMS待定","Assembly","双面以内，含基础测试","服务","装配",18.00,10.00,"不含ICT/FCT治具摊销"],
]

headers=["类别","位号","名称","单板用量","建议品牌","建议型号/料号","关键规格","封装","装配状态","100套单价(元)","1000套单价(元)","1000套单板金额(元)","说明"]

def ext(row):
    return 0 if row[8] == "DNP" else row[3]*row[10]

csv_path=ROOT/'RemoteCAN_Gateway_Full_BOM.csv'
with csv_path.open('w',newline='',encoding='utf-8-sig') as f:
    w=csv.writer(f); w.writerow(headers)
    for r in rows: w.writerow(r[:11]+[round(ext(r),3),r[11]])

wb=Workbook(); ws=wb.active; ws.title='完整BOM'
ws.append(headers)
for r in rows:
    nr=ws.max_row+1
    ws.append(r[:11]+[f'=IF(I{nr}="DNP",0,D{nr}*K{nr})',r[11]])
head_fill=PatternFill('solid',fgColor='1F4E78'); dnp_fill=PatternFill('solid',fgColor='D9EAD3'); warn_fill=PatternFill('solid',fgColor='FFF2CC')
for c in ws[1]: c.font=Font(color='FFFFFF',bold=True); c.fill=head_fill; c.alignment=Alignment(horizontal='center',vertical='center')
for row in ws.iter_rows(min_row=2):
    if row[8].value=='DNP':
        for c in row: c.fill=dnp_fill
    if '待锁定' in str(row[4].value):
        for c in row: c.fill=warn_fill
    for c in row: c.alignment=Alignment(vertical='top',wrap_text=True)
ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:M{ws.max_row}'
widths=[12,18,22,10,24,30,42,18,12,15,16,20,46]
for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
for r in range(2,ws.max_row+1): ws.cell(r,10).number_format='0.000'; ws.cell(r,11).number_format='0.000'; ws.cell(r,12).number_format='0.000'

sumws=wb.create_sheet('成本汇总')
sumws.append(['RemoteCAN Gateway 成本汇总（人民币，规划估算）','数值'])
sumws.append(['1000套：单板/单套物料+制造成本',f'=SUM(完整BOM!L2:L{ws.max_row})'])
sumws.append(['100套：单板/单套物料+制造成本',f'=SUMPRODUCT((完整BOM!I2:I{ws.max_row}<>"DNP")*完整BOM!D2:D{ws.max_row}*完整BOM!J2:J{ws.max_row})'])
sumws.append(['1000套硬件总额', '=B2*1000'])
sumws.append(['说明1','价格为2026-10规划估算，不是供应商正式报价；未含税费、运输、认证、模具、研发和售后。'])
sumws.append(['说明2','DNP默认不计成本；“条件装配”当前计入，量产时根据RF模块电平二选一调整。'])
sumws.append(['说明3','RF模块和遥控器供应商尚未锁定，是当前成本及封装误差的主要风险。'])
for c in sumws[1]: c.font=Font(color='FFFFFF',bold=True); c.fill=head_fill
sumws.column_dimensions['A'].width=48; sumws.column_dimensions['B'].width=70
for row in sumws.iter_rows():
    for c in row: c.alignment=Alignment(vertical='top',wrap_text=True)
sumws['B2'].number_format='0.00'; sumws['B3'].number_format='0.00'; sumws['B4'].number_format='0.00'

wb.save(ROOT/'RemoteCAN_Gateway_Full_BOM.xlsx')
print('rows',len(rows),'1000-unit total',round(sum(ext(r) for r in rows),2))
