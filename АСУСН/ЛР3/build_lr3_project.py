# -*- coding: utf-8 -*-
"""Генерирует конфигурацию проекта Rapid SCADA 6 (lr3) для ЛР3, вариант 1.
Источник данных: modbus_plc_simulator.py (Modbus TCP, 127.0.0.1:502, Slave ID 1)."""
import os, re
from xml.sax.saxutils import escape

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lr3")
NS = 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema"'

def write(rel, text, bom=True):
    path = os.path.join(ROOT, *rel.split("/"))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data = text.replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-8")
    with open(path, "wb") as f:
        f.write((b"\xef\xbb\xbf" if bom else b"") + data)
    print("written", rel)

def read(rel):
    with open(os.path.join(ROOT, *rel.split("/")), encoding="utf-8-sig") as f:
        return f.read()

def nil(tag, val):
    return f"<{tag}>{escape(str(val))}</{tag}>" if val is not None else f'<{tag} xsi:nil="true" />'

# ---------- База конфигурации ----------
write("BaseXML/Obj.xml", f"""<?xml version="1.0" encoding="utf-8"?>
<ArrayOfObj {NS}>
  <Obj>
    <ObjNum>1</ObjNum>
    <Name>Узел подготовки теплоносителя</Name>
    <Code>Lab3</Code>
    <ParentObjNum xsi:nil="true" />
  </Obj>
</ArrayOfObj>""")

write("BaseXML/CommLine.xml", f"""<?xml version="1.0" encoding="utf-8"?>
<ArrayOfCommLine {NS}>
  <CommLine>
    <CommLineNum>1</CommLineNum>
    <Name>Lab3_Modbus</Name>
    <Descr>Modbus TCP, 127.0.0.1:502</Descr>
  </CommLine>
</ArrayOfCommLine>""")

write("BaseXML/Device.xml", f"""<?xml version="1.0" encoding="utf-8"?>
<ArrayOfDevice {NS}>
  <Device>
    <DeviceNum>1</DeviceNum>
    <Name>PLC_Lab3</Name>
    <Code>PLC</Code>
    <DevTypeID>214</DevTypeID>
    <NumAddress>1</NumAddress>
    <StrAddress />
    <CommLineNum>1</CommLineNum>
    <Descr>Виртуальный ПЛК (Modbus TCP Slave, ID 1)</Descr>
  </Device>
</ArrayOfDevice>""")

def lim(i, name, lolo, low, high, hihi, db):
    return f"""  <Lim>
    <LimID>{i}</LimID>
    <Name>{name}</Name>
    <IsBoundToCnl>false</IsBoundToCnl>
    <IsShared>false</IsShared>
    {nil('LoLo', lolo)}
    {nil('Low', low)}
    {nil('High', high)}
    {nil('HiHi', hihi)}
    {nil('Deadband', db)}
  </Lim>"""

write("BaseXML/Lim.xml", f"""<?xml version="1.0" encoding="utf-8"?>
<ArrayOfLim {NS}>
{lim(1, 'Границы температуры (Перегрев: T > 80)', None, 10, 70, 80, 0.5)}
{lim(2, 'Границы давления (Высокое давление: P > 8)', None, None, 7, 8, 0.05)}
{lim(3, 'Границы уровня', 5, 10, 90, 95, 0.5)}
</ArrayOfLim>""")

# размерность "бар"
unit = read("BaseXML/Unit.xml")
if "<UnitID>1002</UnitID>" not in unit:
    unit = unit.replace("</ArrayOfUnit>", """  <Unit>
    <UnitID>1002</UnitID>
    <Name>бар</Name>
    <Code>bar</Code>
    <Descr>Бар</Descr>
  </Unit>
</ArrayOfUnit>""")
    write("BaseXML/Unit.xml", unit)

EV_ALARM = 1 + 2 + 16        # вкл + звук + изменение статуса
EV_STATUS = 1 + 16
EV_CMD = 1 + 64              # вкл + команда

def cnl(num, name, code, ctype, dev, tag, formula, fmt, qty, unit, limid, ev):
    return f"""  <Cnl>
    <CnlNum>{num}</CnlNum>
    <Active>true</Active>
    <Name>{name}</Name>
    <Code>{code}</Code>
    <DataTypeID xsi:nil="true" />
    <DataLen xsi:nil="true" />
    <CnlTypeID>{ctype}</CnlTypeID>
    <ObjNum>1</ObjNum>
    {nil('DeviceNum', dev)}
    <TagNum xsi:nil="true" />
    {'<TagCode>%s</TagCode>' % tag if tag else '<TagCode />'}
    <FormulaEnabled>{'true' if formula else 'false'}</FormulaEnabled>
    {'<InFormula>%s</InFormula>' % escape(formula) if formula else '<InFormula />'}
    <OutFormula />
    {nil('FormatID', fmt)}
    <OutFormatID xsi:nil="true" />
    {nil('QuantityID', qty)}
    {nil('UnitID', unit)}
    {nil('LimID', limid)}
    <ArchiveMask xsi:nil="true" />
    {nil('EventMask', ev)}
  </Cnl>"""

write("BaseXML/Cnl.xml", f"""<?xml version="1.0" encoding="utf-8"?>
<ArrayOfCnl {NS}>
{cnl(1, 'Temp_Value', 'Temp', 1, 1, 'Temp', 'Cnl / 10', 3, 4, 22, 1, EV_ALARM)}
{cnl(2, 'Pressure_Value', 'Pressure', 1, 1, 'Pressure', 'Cnl / 100', 4, 41, 1002, 2, EV_ALARM)}
{cnl(3, 'Level_Value', 'Level', 1, 1, 'Level', 'Cnl / 10', 3, None, 101, 3, EV_STATUS)}
{cnl(4, 'Pump_State', 'PumpState', 1, 1, 'PumpState', None, 21, 11, None, None, EV_STATUS + 8)}
{cnl(5, 'Pump_Control', 'PumpCtrl', 2, 1, 'PumpCtrl', None, 21, 11, None, None, EV_CMD)}
{cnl(6, 'Alarm_Flag', 'Alarm', 3, None, None, '(Val(1) > 80 || Val(2) > 8) ? 1.0 : 0.0', 26, 11, None, None, None)}
</ArrayOfCnl>""")

write("BaseXML/View.xml", f"""<?xml version="1.0" encoding="utf-8"?>
<ArrayOfView {NS}>
  <View>
    <ViewID>1</ViewID>
    <Path>Lab3_Mimic.mim</Path>
    <ViewTypeID>7</ViewTypeID>
    <Title>Lab3_Mimic — мнемосхема узла</Title>
    <Ord>1</Ord>
    <Hidden>false</Hidden>
    <ObjNum>1</ObjNum>
  </View>
  <View>
    <ViewID>2</ViewID>
    <Path>Lab3_Table.tbl</Path>
    <ViewTypeID>1</ViewTypeID>
    <Title>Lab3 — таблица каналов</Title>
    <Ord>2</Ord>
    <Hidden>false</Hidden>
    <ObjNum>1</ObjNum>
  </View>
</ArrayOfView>""")

write("Views/Lab3_Table.tbl", """<?xml version="1.0" encoding="utf-8"?>
<TableView>
  <TableOptions>
    <UseDefault>true</UseDefault>
  </TableOptions>
  <TableItems>
    <Item cnlNum="1" autoText="true" />
    <Item cnlNum="2" autoText="true" />
    <Item cnlNum="3" autoText="true" />
    <Item cnlNum="4" autoText="true" />
    <Item cnlNum="5" autoText="true" />
    <Item cnlNum="6" autoText="true" />
  </TableItems>
</TableView>""")

# ---------- Коммуникатор ----------
comm = read("Instances/Default/ScadaComm/Config/ScadaCommConfig.xml")
lines = """<Lines>
    <Line active="true" isBound="true" number="1" name="Lab3_Modbus">
      <LineOptions>
        <ReqRetries>3</ReqRetries>
        <CycleDelay>1000</CycleDelay>
        <CmdEnabled>true</CmdEnabled>
        <PollAfterCmd>true</PollAfterCmd>
        <DetailedLog>true</DetailedLog>
      </LineOptions>
      <Channel type="TcpClient" driver="DrvCnlBasic">
        <Option name="Behavior" value="Master" />
        <Option name="ConnectionMode" value="Shared" />
        <Option name="DisconnectOnError" value="true" />
        <Option name="Host" value="127.0.0.1" />
        <Option name="ReconnectAfter" value="5" />
        <Option name="StayConnected" value="true" />
        <Option name="TcpPort" value="502" />
      </Channel>
      <CustomOptions>
        <Option name="TransMode" value="TCP" />
      </CustomOptions>
      <DevicePolling>
        <Device active="true" isBound="true" number="1" name="PLC_Lab3" driver="DrvModbus" numAddress="1" strAddress="" pollOnCmd="false" timeout="1000" delay="200" time="00:00:00" period="00:00:00" cmdLine="DrvModbus_Lab3.xml" />
      </DevicePolling>
    </Line>
  </Lines>"""
comm = re.sub(r"<Lines>.*</Lines>", lambda m: lines, comm, flags=re.S)
write("Instances/Default/ScadaComm/Config/ScadaCommConfig.xml", comm, bom=False)

write("Instances/Default/ScadaComm/Config/DrvModbus_Lab3.xml", """<?xml version="1.0" encoding="utf-8"?>
<DeviceTemplate>
  <Options>
    <ZeroAddr>true</ZeroAddr>
    <DecAddr>true</DecAddr>
    <DefByteOrder2 />
    <DefByteOrder4 />
    <DefByteOrder8 />
  </Options>
  <ElemGroups>
    <ElemGroup active="true" dataBlock="HoldingRegisters" address="0" name="Технологические параметры">
      <Elem type="ushort" readOnly="true" tagCode="Temp" name="Температура x10" />
      <Elem type="ushort" readOnly="true" tagCode="Pressure" name="Давление x100" />
      <Elem type="ushort" readOnly="true" tagCode="Level" name="Уровень x10" />
      <Elem type="ushort" readOnly="true" tagCode="PumpState" name="Состояние насоса" />
      <Elem type="ushort" tagCode="PumpCtrl" name="Команда управления насосом" />
    </ElemGroup>
  </ElemGroups>
  <Cmds />
</DeviceTemplate>""", bom=False)

# ---------- Сервер: вернуть минутный архив, оставить Lab3 (1 с, 1 день) ----------
srv = read("Instances/Default/ScadaServer/Config/ScadaServerConfig.xml")
if 'code="Min"' not in srv:
    min_arc = """<Archive active="true" code="Min" name="Minute archive" kind="Historical" module="ModArcBasic">
      <Option name="IsPeriodic" value="true" />
      <Option name="LogEnabled" value="false" />
      <Option name="MaxQueueSize" value="1000" />
      <Option name="PullToPeriod" value="0" />
      <Option name="ReadOnly" value="false" />
      <Option name="Retention" value="365" />
      <Option name="UseCopyDir" value="false" />
      <Option name="WriteWithPeriod" value="true" />
      <Option name="WritingPeriod" value="1" />
      <Option name="WritingPeriodUnit" value="Minute" />
    </Archive>
    """
    srv = srv.replace('<Archive active="true" code="Lab3"', min_arc + '<Archive active="true" code="Lab3"', 1)
    write("Instances/Default/ScadaServer/Config/ScadaServerConfig.xml", srv, bom=False)

# ---------- Вебстанция: графики из архива Lab3, журнал 100 событий ----------
web = read("Instances/Default/ScadaWeb/config/ScadaWebConfig.xml")
web = web.replace('<Option name="ChartArchiveCode" value="Min" />', '<Option name="ChartArchiveCode" value="Lab3" />')
web = re.sub(r'<Option name="GapBetweenPoints" value="\d+" />', '<Option name="GapBetweenPoints" value="10" />', web)
web = re.sub(r'<Option name="EventCount" value="\d+" />', '<Option name="EventCount" value="100" />', web)
web = web.replace("<DefaultStartPage />", "<DefaultStartPage>~/View/1</DefaultStartPage>")
write("Instances/Default/ScadaWeb/config/ScadaWebConfig.xml", web, bom=False)

# ---------- Мнемосхема ----------
_id = [0]
def nid():
    _id[0] += 1
    return _id[0]

def base(x, y, w, h, extra=""):
    return f"""<ID>{nid()}</ID>
      <Enabled>true</Enabled>
      <Visible>true</Visible>
      <Location><X>{x}</X><Y>{y}</Y></Location>
      <Size><Width>{w}</Width><Height>{h}</Height></Size>{extra}"""

def radius(r):
    return f"<CornerRadius><TopLeft>{r}</TopLeft><TopRight>{r}</TopRight><BottomRight>{r}</BottomRight><BottomLeft>{r}</BottomLeft></CornerRadius>"

def panel(x, y, w, h, back, border=None, r=0, extra=""):
    b = f"<Border><Width>{border[0]}</Width><Color>{border[1]}</Color></Border>" if border else ""
    return f"""    <Panel>
      {base(x, y, w, h)}
      <BackColor>{back}</BackColor>{b}{radius(r) if r else ''}{extra}
    </Panel>"""

def text(x, y, w, h, txt, size=14, bold=False, color="#1F2937", align="MiddleLeft", extra=""):
    fc = f"<ForeColor>{color}</ForeColor>" if color else ""
    return f"""    <Text>
      {base(x, y, w, h)}
      <AutoSize>false</AutoSize>
      <Font><Inherit>false</Inherit><Size>{size}</Size><Bold>{'true' if bold else 'false'}</Bold></Font>{fc}
      <TextAlign>{align}</TextAlign>
      <Text>{escape(txt)}</Text>{extra}
    </Text>"""

def conds(items, kind):
    out = '<Conditions isArray="true">'
    for op, arg, val in items:
        out += f"<Item><ComparisonOper1>{op}</ComparisonOper1><ComparisonArg1>{arg}</ComparisonArg1><LogicalOper>None</LogicalOper><ComparisonOper2>None</ComparisonOper2><ComparisonArg2>0</ComparisonArg2><{kind}>{escape(val)}</{kind}></Item>"
    return out + "</Conditions>"

def led(x, y, d, cnl, items, tooltip=""):
    return f"""    <BasicLed>
      {base(x, y, d, d)}
      <InCnlNum>{cnl}</InCnlNum>
      <BackColor>Silver</BackColor>
      <Border><Width>4</Width><Color>Black</Color></Border>
      <BorderOpacity>30</BorderOpacity>
      <IsSquare>false</IsSquare>
      <DefaultColor>Silver</DefaultColor>
      <Tooltip>{escape(tooltip)}</Tooltip>
      {conds(items, 'Color')}
    </BasicLed>"""

def button(x, y, w, h, txt, back, val):
    return f"""    <BasicButton>
      {base(x, y, w, h)}
      <OutCnlNum>5</OutCnlNum>
      <Text>{txt}</Text>
      <Font><Inherit>false</Inherit><Size>20</Size><Bold>true</Bold></Font>
      <BackColor>{back}</BackColor>
      <ForeColor>White</ForeColor>
      {radius(6)}
      <ClickAction>
        <ActionType>SendCommand</ActionType>
        <CommandArgs><ShowDialog>false</ShowDialog><CmdVal>{val}</CmdVal></CommandArgs>
      </ClickAction>
    </BasicButton>"""

CARD = ("#FFFFFF", (1, "#CBD5E1"), 8)
chart = "<ClickAction><ActionType>DrawChart</ActionType></ClickAction><Tooltip>Нажмите, чтобы открыть график</Tooltip>"
TANK_TOP, TANK_H = 162, 200   # внутренняя область резервуара: y = 162..362
level_bind = f"""
      <PropertyBindings isArray="true">
        <Item><PropertyName>size.height</PropertyName><DataSource>3</DataSource><DataMember>Value</DataMember><Expression>isNaN(x) ? 0 : Math.round(Math.max(0, Math.min(100, x)) * {TANK_H // 100})</Expression><Format /></Item>
        <Item><PropertyName>location.y</PropertyName><DataSource>3</DataSource><DataMember>Value</DataMember><Expression>{TANK_TOP + TANK_H} - (isNaN(x) ? 0 : Math.round(Math.max(0, Math.min(100, x)) * {TANK_H // 100}))</Expression><Format /></Item>
      </PropertyBindings>"""

comps = [
    panel(0, 0, 1000, 60, "#1F3A5F"),
    text(20, 10, 960, 40, "АСУ СН — ЛР3. Мнемосхема Lab3_Mimic: узел подготовки теплоносителя (вариант 1)", 20, True, "#FFFFFF", "MiddleCenter"),
    # резервуар
    panel(30, 80, 300, 510, *CARD[:2], r=8),
    text(50, 95, 260, 30, "Резервуар ТН-01", 18, True, align="MiddleCenter"),
    panel(108, 160, 144, 204, "#E5E7EB", (2, "#374151")),
    panel(110, 262, 140, 100, "#2F80ED", extra=level_bind),
    text(260, 150, 60, 24, "100 %", 13, color="#6B7280"),
    text(260, 250, 60, 24, "50 %", 13, color="#6B7280"),
    text(260, 350, 60, 24, "0 %", 13, color="#6B7280"),
    text(50, 385, 260, 26, "Уровень LT-01 (Level_Value)", 14, color="#6B7280", align="MiddleCenter"),
    text(50, 415, 260, 50, "---", 30, True, None, "MiddleCenter", "<InCnlNum>3</InCnlNum>" + chart),
    text(50, 480, 260, 90, "Высота столба жидкости привязана к каналу 3: 0–100 % → 0–200 px", 12, color="#6B7280", align="MiddleCenter", extra="<WordWrap>true</WordWrap>"),
    # температура
    panel(360, 80, 300, 160, *CARD[:2], r=8),
    text(380, 92, 260, 28, "Температура TT-01 (Temp_Value)", 14, color="#6B7280"),
    text(380, 125, 260, 60, "---", 34, True, None, "MiddleLeft", "<InCnlNum>1</InCnlNum>" + chart),
    text(380, 195, 260, 28, "Тревога «Перегрев»: T > 80 °C", 13, color="#6B7280"),
    # давление
    panel(680, 80, 290, 160, *CARD[:2], r=8),
    text(700, 92, 250, 28, "Давление PT-01 (Pressure_Value)", 14, color="#6B7280"),
    text(700, 125, 250, 60, "---", 34, True, None, "MiddleLeft", "<InCnlNum>2</InCnlNum>" + chart),
    text(700, 195, 250, 28, "Тревога «Высокое давление»: P > 8 бар", 13, color="#6B7280"),
    # авария
    panel(360, 260, 610, 140, *CARD[:2], r=8),
    text(380, 272, 570, 28, "Лампа «Авария» (T > 80 °C ИЛИ P > 8 бар)", 14, color="#6B7280"),
    led(390, 312, 64, 6, [("GreaterThan", 0, "Red"), ("LessThanEqual", 0, "LimeGreen")], "Alarm_Flag"),
    text(480, 314, 470, 60, "НЕТ ДАННЫХ", 26, True, "#1F2937", "MiddleLeft",
         "<InCnlNum>6</InCnlNum><DefaultText>НЕТ ДАННЫХ</DefaultText>" +
         conds([("GreaterThan", 0, "АВАРИЯ! Параметры вне нормы"), ("LessThanEqual", 0, "НОРМА")], "Text")),
    # управление
    panel(360, 420, 610, 170, *CARD[:2], r=8),
    text(380, 432, 570, 28, "Управление насосом Н-01 (канал управления Pump_Control)", 14, color="#6B7280"),
    button(390, 480, 170, 70, "ПУСК", "#2E7D32", 1),
    button(580, 480, 170, 70, "СТОП", "#C62828", 0),
    led(780, 492, 44, 4, [("GreaterThan", 0, "LimeGreen"), ("LessThanEqual", 0, "Gray")], "Pump_State"),
    text(835, 485, 125, 60, "нет данных", 15, True, "#1F2937", "MiddleLeft",
         "<InCnlNum>4</InCnlNum><DefaultText>нет данных</DefaultText><WordWrap>true</WordWrap>" +
         conds([("GreaterThan", 0, "Насос работает"), ("LessThanEqual", 0, "Насос остановлен")], "Text")),
]

write("Views/Lab3_Mimic.mim", """<?xml version="1.0" encoding="utf-8"?>
<Mimic editorVersion="6.0.0.0">
  <Dependencies />
  <Document>
    <BackColor>#F1F5F9</BackColor>
    <Font><Inherit>false</Inherit><Name>Arial</Name><Size>14</Size></Font>
    <Size><Width>1000</Width><Height>610</Height></Size>
  </Document>
  <Components>
""" + "\n".join(comps) + """
  </Components>
  <Images />
</Mimic>""")
