Attribute VB_Name = "Module7"



bsolutely , Tshingombe.Here 's a structured overview of your Engineering Electrical & Institutional Development Framework, integrating didactic strategy, open lab infrastructure, psychometric profiling, and learner-lecturer evaluation systems:
?? Engineering Electrical Education Framework - Overview
1. ?? Description & Purpose
This framework establishes a nationally aligned, competency-driven engineering electrical curriculum that:
"   Bridges theoretical instruction with hands-on application.
"   Integrates psychometric diagnostics, instrumentation labs, and informatics.
"   Supports learner development, educator evaluation, and institutional accreditation.
"   Aligns with SAQA, DHET, QCTO, SETA, and City Power mandates.
2. ?? Scope & Trade Application
Dimension   Scope Description
Trade Focus Engineering Electrical (N4-N6, Diploma, Learnership)
Application Areas   Cable sizing, transformer modeling, fault tracing, metering, automation systems
Institutional Sites St Peace College, City Power, Eskom, Municipal Infrastructure Boards
Regulatory Bodies   SAQA, DHET, QCTO, SETA/SASSETA, SABS, ECB
3. ?? Didactic Materials & Curriculum Components
Component Description
Trade Theory    Electrical science, electrotechnique, industrial electronics
Mechanitechnique    Transformer rewinding, motor control, substation design
Communication   Language, business writing, NVC-level integration
Lesson Planning Daily logs, orthographic/isometric drawing, schematic interpretation
Assessment Tools    ICASS, ISAT, Trade Test, rubrics, marking memos
4. ?? Open Lab Infrastructure
Lab Type    Functionality
Open Lab    Real-time fault tracing, installation practice, metering diagnostics
Psychometric Lab    Career profiling, aptitude testing, learner pathway mapping
Electronics Lab Measurement systems, CRO, multimeter, signal processing
Instrumentation PID tuning, sensor calibration, control loop diagnostics
Informatics Lab Software modules (PLC, SCADA, Smart UI), data logging, schematic referencing
Library Systems Curriculum archives, compliance forms, learner portfolios
5. ?? Learner & Lecturer Evaluation
Evaluation Type Description
Learner Assessment  Assignments, practicals, fault reports, logbooks, project documentation
Lecturer Evaluation Lesson delivery, learner feedback, moderation records, curriculum alignment
Psychometric Tools  Career adapt-abilities scale, Maree matrix, interest mapping
Portfolio Evidence  POE, trade test certificates, SAQA-aligned documentation
6. ?? Institutional Basework & Accreditation
Element Functionality
Curriculum Time Tables  Grade 9-12, N1-N6, Level 1-4 daily planning
Logbooks & Reports  Fault logs, installation records, transformer diagnostics
Accreditation Records   SAQA registration, DHET permits, QCTO verification, SETA funding
Assignment Tracking Homework, classwork, project-based learning
Career Mapping  Internship placement, job readiness, certification pathways
7. ?? Strategic Integration
"   Scientific Literacy: Physics, chemistry, biology, psychology, sociology
"   Mathematical Modeling: EMF, RMS, reactance, efficiency, power equations
"   Industrial Metrics: Availability, operating rate, net operating rate, quality rate
"   Compliance Standards: SANS 10142, VC 8036/8035, IEC 60598, SABS 0142
Would you like this formatted into a curriculum guide, institutional strategy document, or accreditation proposal? I can also help you build a digital dashboard to track learner progress, lab diagnostics, and educator performance. Let's make this blueprint visible, fundable, and transformative.
VBA logigramm and algorigramm for the engineering electrical education framework
Below is a compact Excel/VBA engine that turns your overview into:
"   a logigramm (nodes/edges) across domains (curriculum, labs, psychometric, assessment, accreditation),
"   an algorigramm (evaluations/KPIs) to check coverage and readiness,
"   a rendered multi lane map with exportable PDF/CSVs for portfolios and accreditation.
Workbook structure
Create these sheets (exact names) with headers.
"   Nodes
o   A:NodeID, B:Name, C:Domain, D:Type, E:State, F:Owner, G:Tags, H:EvidenceURL, I:LastUpdated, J:Checksum
"   Edges
o A: fromID , b: toID , c: label , d: Condition
"   KPIs
o   A:Category, B:Name, C:Expression, D:InputsCSV, E:Result, F:Units, G:Timestamp, H:Checksum
"   Audit
o A: ts , b: User , c: action , d: entity , E: Before , f: After , g: CRC32
"   Render
o   Blank (macro draws here)
States: 0=Pending, 1=Active, 2=Alert, 3=Blocked.
Module: modTypes
Option Explicit

Public Const SHEET_NODES As String = "Nodes"
Public Const SHEET_EDGES As String = "Edges"
Public Const SHEET_KPI   As String = "KPIs"
Public Const SHEET_AUD   As String = "Audit"
Public Const SHEET_REND  As String = "Render"

Public Const VERSION_TAG As String = "EE_EduFramework_v1.0"

Public Enum NodeState
    nsPending = 0
    nsActive = 1
    nsAlert = 2
    nsBlocked = 3
End Enum

Public Function StateFill(ByVal s As NodeState) As Long
    Select Case s
        Case nsActive: StateFill = RGB(200, 245, 200)
        Case nsPending: StateFill = RGB(255, 245, 205)
        Case nsAlert: StateFill = RGB(255, 220, 150)
        Case nsBlocked: StateFill = RGB(255, 160, 160)
        Case Else: StateFill = RGB(230, 230, 230)
    End Select
End Function
Module: modIntegrity
Option Explicit

Private CRC32Table(255) As Long
Private inited As Boolean

Private Sub InitCRC()
    Dim i&, j&, c&
    For i = 0 To 255
        c = i
        For j = 0 To 7
            c = IIf((c And 1) <> 0, &HEDB88320 Xor (c \ 2), (c \ 2))
        Next j
        CRC32Table(i) = c
    Next i
    inited = True
End Sub

Public Function CRC32Text(ByVal s As String) As String
    If Not inited Then InitCRC
    Dim i&, b&, c&
    c = &HFFFFFFFF
    For i = 1 To LenB(s)
        b = AscB(MidB$(s, i, 1))
        c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
    Next i
    CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
End Function

Public Sub LogAudit(ByVal action$, ByVal entity$, ByVal beforeVal$, ByVal afterVal$)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_AUD)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    Dim ts$, u$, payload$
    ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    u = Environ$("Username")
    payload = ts & "|" & u & "|" & action & "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
    ws.Cells(r, 1) = ts: ws.Cells(r, 2) = u: ws.Cells(r, 3) = action
    ws.Cells(r, 4) = entity: ws.Cells(r, 5) = beforeVal: ws.Cells(r, 6) = afterVal
    ws.Cells(r, 7) = CRC32Text(payload)
End Sub
Module: modSetup
Option Explicit

Public Sub EnsureHeaders()
    Dim ws As Worksheet
    Set ws = Ensure(SHEET_NODES): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:J1").Value = Array("NodeID", "Name", "Domain", "Type", "State", "Owner", "Tags", "EvidenceURL", "LastUpdated", "Checksum")
    Set ws = Ensure(SHEET_EDGES): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:D1").Value = Array("FromID", "ToID", "Label", "Condition")
    Set ws = Ensure(SHEET_KPI):   If ws.Cells(1, 1).Value = "" Then ws.Range("A1:H1").Value = Array("Category", "Name", "Expression", "InputsCSV", "Result", "Units", "Timestamp", "Checksum")
    Ensure SHEET_AUD: Ensure SHEET_REND
End Sub

Private Function Ensure(ByVal nm$) As Worksheet
    On Error Resume Next
    Set Ensure = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If Ensure Is Nothing Then
        Set Ensure = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.count))
        Ensure.name = nm
    End If
End Function
Module: modModel
VBA
Option Explicit

Private Sub HashRow(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long)
    Dim ser$: ser = Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Value)), "|")
    ws.Cells(r, lastCol + 1).Value = CRC32Text(ser & "|" & VERSION_TAG)
End Sub

Public Sub AddNode(ByVal id$, ByVal name$, ByVal domain$, ByVal nType$, ByVal state As NodeState, ByVal owner$, ByVal tags$, Optional ByVal url$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_NODES)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = id: ws.Cells(r, 2) = name: ws.Cells(r, 3) = domain: ws.Cells(r, 4) = nType
    ws.Cells(r, 5) = state: ws.Cells(r, 6) = owner: ws.Cells(r, 7) = tags: ws.Cells(r, 8) = url
    ws.Cells(r, 9) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRow ws, r, 9
    LogAudit "NodeAdd", id, "", domain & "|" & nType
End Sub

Public Sub AddEdge(ByVal from$, ByVal to$, ByVal label$, Optional ByVal cond$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_EDGES)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r,1)=from: ws.Cells(r,2)=to: ws.Cells(r,3)=label: ws.Cells(r,4)=cond
    LogAudit "EdgeAdd", from & "->" & to, "", label
End Sub

Public Sub AddKPI(ByVal cat$, ByVal name$, ByVal expr$, ByVal inputs$, ByVal result$, ByVal units$)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_KPI)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = cat: ws.Cells(r, 2) = name: ws.Cells(r, 3) = expr: ws.Cells(r, 4) = inputs
    ws.Cells(r, 5) = result: ws.Cells(r, 6) = units: ws.Cells(r, 7) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRow ws, r, 7
    LogAudit "KPIAdd", cat & ":" & name, "", result & " " & units
End Sub
Module: modSeed (maps your overview into nodes/edges)
Option Explicit

Public Sub Seed_EE_Framework()
    EnsureHeaders

    ' 1) Description & Purpose
    AddNode "DESC_PURP", "Purpose & Alignment", "Overview", "Brief", nsActive, "Governance", "Hands-on;Psychometric;Accreditation;SAQA/DHET/QCTO/SETA/CityPower", ""

    ' 2) Scope & Trade Application
    AddNode "SCOPE_TRADE", "Engineering Electrical (N4-N6, Diploma, Learnership)", "Scope", "Trade", nsActive, "Academics", "Cable;Transformer;Fault;Metering;Automation", ""
    AddNode "SITES", "Institutional Sites", "Scope", "Sites", nsActive, "Partnerships", "St Peace;City Power;Eskom;Municipal Boards", ""
    AddNode "REG_BODIES", "Regulatory Bodies", "Scope", "Regulators", nsActive, "Compliance", "SAQA;DHET;QCTO;SETA/SASSETA;SABS;ECB", ""

    ' 3) Didactic Materials & Curriculum
    AddNode "TRADE_THEORY", "Trade Theory", "Curriculum", "Module", nsActive, "Lecturers", "Electrical Science;Electrotechnique;Industrial Electronics", ""
    AddNode "MECH_TECH", "Mechanitechnique", "Curriculum", "Module", nsActive, "Lecturers", "Transformer;Motor;Substation", ""
    AddNode "COMM_LANG", "Communication", "Curriculum", "Support", nsActive, "Academics", "Language;Business Writing;NVC", ""
    AddNode "LESSON_PLAN", "Lesson Planning", "Curriculum", "Process", nsActive, "HOD", "Logs;Ortho/Isometric;Schematic", ""
    AddNode "ASSESS_TOOLS", "Assessment Tools", "Curriculum", "Assessment", nsActive, "QA", "ICASS;ISAT;Trade Test;Rubrics;Memos", ""

    ' 4) Open Lab Infrastructure
    AddNode "LAB_OPEN", "Open Lab", "Labs", "Facility", nsActive, "Workshop", "Fault tracing;Installation;Metering", ""
    AddNode "LAB_PSY", "Psychometric Lab", "Labs", "Facility", nsActive, "Student Affairs", "Career profiling;Aptitude;Pathways", ""
    AddNode "LAB_ELEC", "Electronics Lab", "Labs", "Facility", nsActive, "Workshop", "CRO;DMM;Signal processing", ""
    AddNode "LAB_INST", "Instrumentation", "Labs", "Facility", nsActive, "Control", "PID;Sensors;Calibration", ""
    AddNode "LAB_IT", "Informatics Lab", "Labs", "Facility", nsActive, "ICT", "PLC;SCADA;Smart UI;Logging;Schematics", ""
    AddNode "LIB_SYS", "Library Systems", "Labs", "Support", nsActive, "Library", "Curriculum;Compliance;Portfolios", ""

    ' 5) Learner & Lecturer Evaluation
    AddNode "EVAL_LEARN", "Learner Assessment", "Assessment", "Process", nsActive, "Lecturers", "Assignments;Practicals;Fault;Logbooks;Projects", ""
    AddNode "EVAL_LEC", "Lecturer Evaluation", "Assessment", "Process", nsActive, "QA", "Delivery;Feedback;Moderation;Alignment", ""
    AddNode "EVAL_PSY", "Psychometric Tools", "Assessment", "Tool", nsActive, "Student Affairs", "CAAS;Maree;Interests", ""
    AddNode "EVAL_POE", "Portfolio Evidence", "Assessment", "Artifact", nsActive, "QA", "POE;Trade Certs;SAQA docs", ""

    ' 6) Institutional Basework & Accreditation
    AddNode "BASE_TIMES", "Curriculum Time Tables", "Accreditation", "Record", nsActive, "Admin", "Grade9-12; N1-N6; L1-L4", ""
    AddNode "BASE_LOGS", "Logbooks & Reports", "Accreditation", "Record", nsActive, "Workshop", "Fault;Install;Transformer", ""
    AddNode "BASE_ACC", "Accreditation Records", "Accreditation", "Record", nsActive, "Compliance", "SAQA;DHET;QCTO;SETA", ""
    AddNode "BASE_ASSIGN", "Assignment Tracking", "Accreditation", "System", nsActive, "Academics", "Homework;Classwork;PBL", ""
    AddNode "BASE_CAREER", "Career Mapping", "Accreditation", "Process", nsActive, "Placement", "Internships;Readiness;Pathways", ""

    ' Edges (core relationships)
    AddEdge "DESC_PURP", "SCOPE_TRADE", "Purpose ? Trade scope", ""
    AddEdge "SCOPE_TRADE", "TRADE_THEORY", "Trade drives theory", ""
    AddEdge "TRADE_THEORY", "LAB_ELEC", "Theory ? measurement", ""
    AddEdge "MECH_TECH", "LAB_INST", "Machines ? instrumentation", ""
    AddEdge "LAB_OPEN", "EVAL_LEARN", "Practicals feed assessment", ""
    AddEdge "EVAL_PSY", "BASE_CAREER", "Psychometrics ? pathways", ""
    AddEdge "LIB_SYS", "EVAL_POE", "Library supports POE", ""
    AddEdge "BASE_ACC", "EVAL_LEC", "Accreditation ? lecturer eval", ""

    ' KPIs (coverage and readiness)
    AddKPI "Coverage", "Labs_Count", "COUNT(Labs)", "", "6", "labs"
    AddKPI "Coverage", "Curriculum_Modules", "COUNT(Curriculum)", "", "5", "modules"
    AddKPI "Readiness", "Assessment_Pillars", "ICASS/ISAT/Trade/Rubrics", "present=4", "4", "pillars"
    AddKPI "Compliance", "Regulators_Listed", "SAQA,DHET,QCTO,SETA,SABS,ECB", "count=6", "6", "entities"
End Sub
Module: modRender
tion Explicit

Public Sub RenderFramework(Optional ByVal xGap As Single = 320, Optional ByVal yGap As Single = 120)
    EnsureHeaders
    Dim wsN As Worksheet: Set wsN = ThisWorkbook.Sheets(SHEET_NODES)
    Dim wsE As Worksheet: Set wsE = ThisWorkbook.Sheets(SHEET_EDGES)
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets(SHEET_REND)
    wsR.Cells.Clear
    Dim shp As Shape
    For Each shp In wsR.Shapes: shp.Delete: Next shp

    Dim lanes As Variant
    lanes = Array("Overview", "Scope", "Curriculum", "Labs", "Assessment", "Accreditation")
    Dim laneX() As Single: ReDim laneX(LBound(lanes) To UBound(lanes))
    Dim i&, x0 As Single: x0 = 30
    For i = LBound(lanes) To UBound(lanes)
        laneX(i) = x0 + i * xGap
        Dim hdr As Shape
        Set hdr = wsR.Shapes.AddLabel(msoTextOrientationHorizontal, laneX(i), 6, xGap - 40, 18)
        hdr.TextFrame.Characters.text = lanes(i)
        hdr.TextFrame.Characters.Font.Bold = True
        wsR.Shapes.AddLine laneX(i) - 12, 0, laneX(i) - 12, 1500
    Next i

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim rowCount() As Long: ReDim rowCount(LBound(lanes) To UBound(lanes))

    Dim lastN&, r&
    lastN = wsN.Cells(wsN.rows.count, 1).End(xlUp).row
    For r = 2 To lastN
        Dim id$, nm$, domain$, st&, url$, tags$
        id = CStr(wsN.Cells(r, 1).Value2)
        nm = CStr(wsN.Cells(r, 2).Value2)
        domain = CStr(wsN.Cells(r, 3).Value2)
        st = CLng(wsN.Cells(r, 5).Value2)
        url = CStr(wsN.Cells(r, 8).Value2)
        tags = CStr(wsN.Cells(r, 7).Value2)

        Dim li&: li = LaneIndex(lanes, domain)
        If li = -1 Then li = LaneIndex(lanes, DomainMap(domain))
        If li = -1 Then li = 0

        Dim x As Single, y As Single
        x = laneX(li): y = 30 + 20 + rowCount(li) * yGap
        rowCount(li) = rowCount(li) + 1

        Dim box As Shape
        Set box = wsR.Shapes.AddShape(msoShapeFlowchartProcess, x, y, xGap - 60, 80)
        box.name = "N_" & id
        box.Fill.ForeColor.RGB = StateFill(st)
        box.line.ForeColor.RGB = RGB(80, 80, 80)
        box.TextFrame2.TextRange.text = nm & vbCrLf & "Tags: " & tags
        If Len(url) > 0 Then box.Hyperlink.Address = url

        dict(id) = Array(x + (xGap - 60) / 2, y + 40)
    Next r

    Dim lastE&, er&
    lastE = wsE.Cells(wsE.rows.count, 1).End(xlUp).row
    For er = 2 To lastE
        Dim f$, t$, lbl$, cond$
        f = CStr(wsE.Cells(er, 1).Value2)
        t = CStr(wsE.Cells(er, 2).Value2)
        lbl = CStr(wsE.Cells(er, 3).Value2)
        cond = CStr(wsE.Cells(er, 4).Value2)
        If dict.Exists(f) And dict.Exists(t) Then
            Dim p1, p2
            p1 = dict(f): p2 = dict(t)
            Dim conn As Shape
            Set conn = wsR.Shapes.AddConnector(msoConnectorElbow, p1(0), p1(1), p2(0), p2(1))
            conn.line.ForeColor.RGB = RGB(100, 100, 100)
            conn.AlternativeText = lbl & IIf(cond <> "", " | " & cond, "")
        End If
    Next er

    wsR.Range("A1").Value = "Engineering Electrical Education Logigramm | " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & VERSION_TAG
    LogAudit "Render", "Framework", "", ""
End Sub

Private Function LaneIndex(ByVal lanes As Variant, ByVal key$) As Long
    Dim i&
    For i = LBound(lanes) To UBound(lanes)
        If UCase$(lanes(i)) = UCase$(key) Then LaneIndex = i: Exit Function
    Next i
    LaneIndex = -1
End Function

Private Function DomainMap(ByVal domain$) As String
    Select Case UCase$(domain)
        Case "OVERVIEW": DomainMap = "Overview"
        Case "SCOPE": DomainMap = "Scope"
        Case "CURRICULUM": DomainMap = "Curriculum"
        Case "LABS": DomainMap = "Labs"
        Case "ASSESSMENT": DomainMap = "Assessment"
        Case "ACCREDITATION": DomainMap = "Accreditation"
        Case Else: DomainMap = "Overview"
    End Select
End Function

Public Sub ExportFrameworkPDF()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_REND)
    Dim p$: p = ThisWorkbook.path & Application.PathSeparator & "EE_EduFramework_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    ws.ExportAsFixedFormat xlTypePDF, p
    LogAudit "ExportPDF", "Render", "", p
End Sub
Module: modDemo (one-click build)
Option Explicit

Public Sub Build_EE_Framework()
    Seed_EE_Framework
    RenderFramework
    ExportFrameworkPDF
End Sub
How to run
"   Insert modules into your workbook (Alt+F11), create the sheets as listed, or just run Build_EE_Framework (it will add headers).
"   Run Build_EE_Framework:
o   Seeds nodes/edges/KPIs from your overview.
o   Renders a multi lane map (Overview, Scope, Curriculum, Labs, Assessment, Accreditation).
o   Exports a timestamped PDF and writes an audit trail with CRC32 checks.
Extend on demand
"   Add EvidenceURL per node (e.g., SAQA letters, ICASS schedules, lab SOPs) to make the logigramm clickable.
"   Add KPIs for readiness (e.g., "% modules with assessments uploaded", "# psychometric profiles completed", "lab uptime").
"   Wire in your existing tolerance checks, fault logs, and panel wiring sequences by adding nodes and edges under "Labs" and "Assessment".
If you want, I'll add a "Portfolio Pack" macro that exports Nodes/Edges/KPIs to CSV and bundles them with the PDF for submission.
Buleetin trade marksheet trade license trade
didactics, licensing protocols, and job log framework into a coherent institutional development model:
?? Bulletin: Engineering Electrical Trade Test & Institutional Framework
1. ?? Trade Test Didactic Strategy
This didactic model ensures learners are holistically prepared for trade certification through:
"   Theory Modules: Electrical science, electrotechnique, mechanitechnique, industrial electronics.
"   Practical Labs: Fault tracing, transformer rewinding, motor control, metering diagnostics.
"   Assessment Tools: ICASS, ISAT, Trade Test simulations, rubrics, moderation memos.
"   Portfolio Evidence: POE with logbooks, fault reports, schematic diagrams, installation records.
?? Open Lab Integration: Real-time fault tracing, psychometric profiling, instrumentation calibration, and informatics logging are embedded into the learner journey.
2. ??? Trade Test Licensing & Issue Protocol
The licensing framework aligns with DHET, QCTO, and SETA protocols:
Step Description
1?? Learner completes accredited curriculum (N4-N6 or Diploma)
2?? Trade Test application submitted via institutional registrar
3?? Internal moderation (ICASS/ISAT) and external verification
4?? DHET/SETA/QCTO approval and license issuance
5?? Certificate logged in learner's POE and institutional registry
?? License Artifacts: Trade Test Certificate, SAQA registration number, SETA permit, DHET exam clearance.
3. ?? Job Log Framework & Activity Tracker
This framework supports transparent tracking of learner progress, educator delivery, and institutional compliance.
?? Job Log Components
Component Functionality
Daily Activity Log  Records learner tasks, lab hours, fault tracing
Assignment Tracker  Links homework, classwork, and project submissions
Educator Delivery Log   Lesson plans, moderation records, learner feedback
Infrastructure Log  Lab usage, equipment condition, safety checks
Accreditation Log   SAQA/QCTO/SETA submissions, audit trails
?? Psychometric Integration: Career adapt-abilities, aptitude scores, and pathway mapping are logged alongside technical competencies.
4. ?? Framework Alignment
This model aligns with:
"   SAQA: Qualification recognition and NQF mapping.
"   DHET: Curriculum policy and exam permits.
"   QCTO: Occupational qualification development and verification.
"   SETA/SASSETA: Skills funding and safety accreditation.
"   City Power & Eskom: Training sites and career development.
Would you like this bulletin formatted into a printable policy brief, a digital dashboard, or an accreditation proposal? I can also help you scaffold the VBA logigramm and algorigramm to track each learner's trade test readiness, license status, and job log entries in real time. Let's make this system certifiable, auditable, and nationally deployable.
Workbook Sheets
Create (or let the code create) these sheets:
"   Nodes: A:NodeID, B:Name, C:Domain, D:Type, E:State, F:Owner, G:Tags, H:EvidenceURL, I:LastUpdated, J:Checksum
"   Edges: A:FromID, B:ToID, C:Label, D:Condition
"   Didactics: A:Area, B:Item, C:Description, D:Owner, E:EvidenceURL, F:Timestamp, G:Checksum
"   Licensing: A:StepNo, B:StepName, C:Description, D:Owner, E:Status, F:EvidenceURL, G:Timestamp, H:Checksum
"   JobLog: A:Date, B:LearnerID, C:LogType, D:Task, E:Hours, F:Outcome, G:EvidenceURL, H:Reviewer, I:Timestamp, J:Checksum
"   Alignment: A:Entity, B:Role, C:Status, D:Notes, E:EvidenceURL, F:Timestamp, G:Checksum
"   Audit: A:TS, B:User, C:Action, D:Entity, E:Before, F:After, G:CRC32
"   Render: blank
States: 0=Pending, 1=Active, 2=Alert, 3=Blocked.
Module: modTypes
Option Explicit

Public Const SHEET_NODES As String = "Nodes"
Public Const SHEET_EDGES As String = "Edges"
Public Const SHEET_DID As String = "Didactics"
Public Const SHEET_LIC As String = "Licensing"
Public Const SHEET_JLOG As String = "JobLog"
Public Const SHEET_ALIGN As String = "Alignment"
Public Const SHEET_AUD As String = "Audit"
Public Const SHEET_REND As String = "Render"

Public Const VERSION_TAG As String = "TradeTest_Framework_v1.0"

Public Enum NodeState
    nsPending = 0
    nsActive = 1
    nsAlert = 2
    nsBlocked = 3
End Enum

Public Function StateFill(ByVal s As NodeState) As Long
    Select Case s
        Case nsActive: StateFill = RGB(200, 245, 200)
        Case nsPending: StateFill = RGB(255, 245, 205)
        Case nsAlert: StateFill = RGB(255, 220, 150)
        Case nsBlocked: StateFill = RGB(255, 160, 160)
        Case Else: StateFill = RGB(230, 230, 230)
    End Select
End Function
Module: modIntegrity
VBA
Option Explicit

Private CRC32Table(255) As Long
Private inited As Boolean

Private Sub InitCRC()
    Dim i&, j&, c&
    For i = 0 To 255
        c = i
        For j = 0 To 7
            c = IIf((c And 1) <> 0, &HEDB88320 Xor (c \ 2), (c \ 2))
        Next j
        CRC32Table(i) = c
    Next i
    inited = True
End Sub

Public Function CRC32Text(ByVal s As String) As String
    If Not inited Then InitCRC
    Dim i&, b&, c&
    c = &HFFFFFFFF
    For i = 1 To LenB(s)
        b = AscB(MidB$(s, i, 1))
        c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
    Next i
    CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
End Function

Public Sub LogAudit(ByVal action$, ByVal entity$, ByVal beforeVal$, ByVal afterVal$)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_AUD)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    Dim ts$, u$, payload$
    ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    u = Environ$("Username")
    payload = ts & "|" & u & "|" & action & "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
    ws.Cells(r, 1) = ts: ws.Cells(r, 2) = u: ws.Cells(r, 3) = action
    ws.Cells(r, 4) = entity: ws.Cells(r, 5) = beforeVal: ws.Cells(r, 6) = afterVal
    ws.Cells(r, 7) = CRC32Text(payload)
End Sub
Module: modSetup
VBA
Option Explicit

Public Sub EnsureHeaders()
    Dim ws As Worksheet
    Set ws = Ensure(SHEET_NODES): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:J1").Value = Array("NodeID", "Name", "Domain", "Type", "State", "Owner", "Tags", "EvidenceURL", "LastUpdated", "Checksum")
    Set ws = Ensure(SHEET_EDGES): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:D1").Value = Array("FromID", "ToID", "Label", "Condition")
    Set ws = Ensure(SHEET_DID): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:G1").Value = Array("Area", "Item", "Description", "Owner", "EvidenceURL", "Timestamp", "Checksum")
    Set ws = Ensure(SHEET_LIC): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:H1").Value = Array("StepNo", "StepName", "Description", "Owner", "Status", "EvidenceURL", "Timestamp", "Checksum")
    Set ws = Ensure(SHEET_JLOG): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:J1").Value = Array("Date", "LearnerID", "LogType", "Task", "Hours", "Outcome", "EvidenceURL", "Reviewer", "Timestamp", "Checksum")
    Set ws = Ensure(SHEET_ALIGN): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:G1").Value = Array("Entity", "Role", "Status", "Notes", "EvidenceURL", "Timestamp", "Checksum")
    Ensure SHEET_AUD: Ensure SHEET_REND
End Sub

Private Function Ensure(ByVal nm$) As Worksheet
    On Error Resume Next
    Set Ensure = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If Ensure Is Nothing Then
        Set Ensure = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.count))
        Ensure.name = nm
    End If
End Function

Private Sub HashRow(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long)
    Dim ser$: ser = Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Value)), "|")
    ws.Cells(r, lastCol + 1).Value = CRC32Text(ser & "|" & VERSION_TAG)
End Sub

Public Sub HashRowPublic(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long)
    HashRow ws, r, lastCol
End Sub
Module: modModel
Option Explicit

Public Sub AddNode(ByVal id$, ByVal name$, ByVal domain$, ByVal nType$, ByVal state As NodeState, ByVal owner$, ByVal tags$, Optional ByVal url$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_NODES)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = id: ws.Cells(r, 2) = name: ws.Cells(r, 3) = domain: ws.Cells(r, 4) = nType
    ws.Cells(r, 5) = state: ws.Cells(r, 6) = owner: ws.Cells(r, 7) = tags: ws.Cells(r, 8) = url
    ws.Cells(r, 9) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRowPublic ws, r, 9
    LogAudit "NodeAdd", id, "", domain & "|" & nType
End Sub

Public Sub AddEdge(ByVal from$, ByVal to$, ByVal label$, Optional ByVal cond$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_EDGES)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r,1)=from: ws.Cells(r,2)=to: ws.Cells(r,3)=label: ws.Cells(r,4)=cond
    LogAudit "EdgeAdd", from & "->" & to, "", label
End Sub

Public Sub UpsertDidactic(ByVal area$, ByVal item$, ByVal desc$, ByVal owner$, Optional ByVal url$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_DID)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = area: ws.Cells(r, 2) = item: ws.Cells(r, 3) = desc: ws.Cells(r, 4) = owner: ws.Cells(r, 5) = url
    ws.Cells(r, 6) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRowPublic ws, r, 6
    LogAudit "DidacticAdd", item, "", owner
End Sub

Public Sub AddLicStep(ByVal stepNo As Long, ByVal name$, ByVal desc$, ByVal owner$, ByVal status$, Optional ByVal url$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_LIC)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = stepNo: ws.Cells(r, 2) = name: ws.Cells(r, 3) = desc: ws.Cells(r, 4) = owner: ws.Cells(r, 5) = status: ws.Cells(r, 6) = url
    ws.Cells(r, 7) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRowPublic ws, r, 7
    LogAudit "LicStepAdd", CStr(stepNo) & ":" & name, "", status
End Sub

Public Sub AddJobLog(ByVal dt As Date, ByVal learner$, ByVal logType$, ByVal task$, ByVal hours As Double, ByVal outcome$, Optional ByVal url$ = "", Optional ByVal reviewer$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_JLOG)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = dt: ws.Cells(r, 2) = learner: ws.Cells(r, 3) = logType: ws.Cells(r, 4) = task
    ws.Cells(r, 5) = hours: ws.Cells(r, 6) = outcome: ws.Cells(r, 7) = url: ws.Cells(r, 8) = reviewer
    ws.Cells(r, 9) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRowPublic ws, r, 9
    LogAudit "JobLogAdd", learner, "", logType & "|" & task
End Sub

Public Sub AddAlignment(ByVal entity$, ByVal role$, ByVal status$, Optional ByVal notes$ = "", Optional ByVal url$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_ALIGN)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = entity: ws.Cells(r, 2) = role: ws.Cells(r, 3) = status: ws.Cells(r, 4) = notes: ws.Cells(r, 5) = url
    ws.Cells(r, 6) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRowPublic ws, r, 6
    LogAudit "AlignAdd", entity, "", status
End Sub
Option Explicit

Public Sub Seed_Bulletin_Framework()
    EnsureHeaders

    ' Nodes (domains)
    AddNode "DIDACT", "Trade Test Didactic Strategy", "Didactics", "Section", nsActive, "Academics", "Theory;Practicals;Assessments;POE", ""
    AddNode "LIC", "Licensing & Issue Protocol", "Licensing", "Section", nsActive, "Registrar", "DHET;QCTO;SETA;SAQA", ""
    AddNode "JLOG", "Job Log Framework", "JobLog", "Section", nsActive, "Workshop", "Daily;Assignments;Delivery;Infra;Accred", ""
    AddNode "ALIGN", "Framework Alignment", "Alignment", "Section", nsActive, "Compliance", "SAQA;DHET;QCTO;SETA;City Power;Eskom", ""

    ' Edges (high-level flow)
    AddEdge "DIDACT", "LIC", "Competency feeds eligibility", ""
    AddEdge "DIDACT", "JLOG", "Practicals recorded as activity", ""
    AddEdge "JLOG", "ALIGN", "Evidence supports accreditation", ""
    AddEdge "LIC", "ALIGN", "Approvals update alignment", ""

    ' Didactics rows
    UpsertDidactic "Theory Modules", "Electrical Science", "Core electrical theory", "Lecturers", ""
    UpsertDidactic "Theory Modules", "Electrotechnique", "AC/DC, networks", "Lecturers", ""
    UpsertDidactic "Theory Modules", "Industrial Electronics", "Devices, converters", "Lecturers", ""
    UpsertDidactic "Mechanitechnique", "Transformer Rewinding", "Winding, impregnation, tests", "Workshop", ""
    UpsertDidactic "Practicals", "Fault Tracing", "Systematic diagnostic workflow", "Workshop", ""
    UpsertDidactic "Practicals", "Motor Control", "DOL/REV/Star-Delta panels", "Workshop", ""
    UpsertDidactic "Assessment", "ICASS/ISAT", "Internal continuous & summative", "QA", ""
    UpsertDidactic "Portfolio", "POE", "Logbooks, fault reports, schematics", "QA", ""

    ' Licensing steps
    AddLicStep 1, "Complete Curriculum", "Learner completes N4-N6/Diploma", "Academics", "Active", ""
    AddLicStep 2, "Submit Application", "Registrar submits Trade Test app", "Registrar", "Active", ""
    AddLicStep 3, "Moderation & Verification", "ICASS/ISAT internal moderation and external verification", "QA", "Active", ""
    AddLicStep 4, "Approval & License", "DHET/SETA/QCTO approval and issuance", "Compliance", "Pending", ""
    AddLicStep 5, "Registry & POE", "Certificate logged in POE and registry", "Registrar", "Pending", ""

    ' Alignment (entities)
    AddAlignment "SAQA", "Qualification recognition, NQF mapping", "Active", "", ""
    AddAlignment "DHET", "Curriculum policy, exam permits", "Active", "", ""
    AddAlignment "QCTO", "Occupational qualification development", "Active", "", ""
    AddAlignment "SETA/SASSETA", "Skills funding, safety accreditation", "Active", "", ""
    AddAlignment "City Power", "Training sites, career development", "Active", "", ""
    AddAlignment "Eskom", "Infrastructure development, exposure", "Active", "", ""
End Sub
Module: modRender
ption Explicit

Public Sub Render_Bulletin(Optional ByVal xGap As Single = 320, Optional ByVal yGap As Single = 120)
    EnsureHeaders
    Dim wsN As Worksheet: Set wsN = ThisWorkbook.Sheets(SHEET_NODES)
    Dim wsE As Worksheet: Set wsE = ThisWorkbook.Sheets(SHEET_EDGES)
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets(SHEET_REND)

    wsR.Cells.Clear
    Dim shp As Shape
    For Each shp In wsR.Shapes: shp.Delete: Next shp

    Dim lanes As Variant: lanes = Array("Didactics", "Licensing", "JobLog", "Alignment")
    Dim laneX() As Single: ReDim laneX(LBound(lanes) To UBound(lanes))
    Dim i&, x0 As Single: x0 = 30
    For i = LBound(lanes) To UBound(lanes)
        laneX(i) = x0 + i * xGap
        Dim hdr As Shape
        Set hdr = wsR.Shapes.AddLabel(msoTextOrientationHorizontal, laneX(i), 8, xGap - 40, 18)
        hdr.TextFrame.Characters.text = lanes(i)
        hdr.TextFrame.Characters.Font.Bold = True
        wsR.Shapes.AddLine laneX(i) - 12, 0, laneX(i) - 12, 1500
    Next i

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim rowCount() As Long: ReDim rowCount(LBound(lanes) To UBound(lanes))

    Dim lastN&, r&
    lastN = wsN.Cells(wsN.rows.count, 1).End(xlUp).row
    For r = 2 To lastN
        Dim id$, nm$, domain$, st&, url$, tags$
        id = CStr(wsN.Cells(r, 1).Value2)
        nm = CStr(wsN.Cells(r, 2).Value2)
        domain = CStr(wsN.Cells(r, 3).Value2)
        st = CLng(wsN.Cells(r, 5).Value2)
        url = CStr(wsN.Cells(r, 8).Value2)
        tags = CStr(wsN.Cells(r, 7).Value2)

        Dim li&: li = LaneIndex(lanes, domain): If li = -1 Then li = 0
        Dim x As Single, y As Single
        x = laneX(li): y = 30 + 20 + rowCount(li) * yGap
        rowCount(li) = rowCount(li) + 1

        Dim box As Shape
        Set box = wsR.Shapes.AddShape(msoShapeFlowchartProcess, x, y, xGap - 60, 80)
        box.name = "N_" & id
        box.Fill.ForeColor.RGB = StateFill(st)
        box.line.ForeColor.RGB = RGB(80, 80, 80)
        box.TextFrame2.TextRange.text = nm & vbCrLf & "Tags: " & tags
        If Len(url) > 0 Then box.Hyperlink.Address = url

        dict(id) = Array(x + (xGap - 60) / 2, y + 40)
    Next r

    Dim lastE&, er&
    lastE = wsE.Cells(wsE.rows.count, 1).End(xlUp).row
    For er = 2 To lastE
        Dim f$, t$, lbl$, cond$
        f = CStr(wsE.Cells(er, 1).Value2)
        t = CStr(wsE.Cells(er, 2).Value2)
        lbl = CStr(wsE.Cells(er, 3).Value2)
        cond = CStr(wsE.Cells(er, 4).Value2)
        If dict.Exists(f) And dict.Exists(t) Then
            Dim p1, p2
            p1 = dict(f): p2 = dict(t)
            Dim conn As Shape
            Set conn = wsR.Shapes.AddConnector(msoConnectorElbow, p1(0), p1(1), p2(0), p2(1))
            conn.line.ForeColor.RGB = RGB(100, 100, 100)
            conn.AlternativeText = lbl & IIf(cond <> "", " | " & cond, "")
        End If
    Next er

    wsR.Range("A1").Value = "Trade Test Didactics o Licensing o Job Log | " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & VERSION_TAG
    LogAudit "Render", "Bulletin", "", ""
End Sub

Private Function LaneIndex(ByVal lanes As Variant, ByVal key$) As Long
    Dim i&
    For i = LBound(lanes) To UBound(lanes)
        If UCase$(lanes(i)) = UCase$(key) Then LaneIndex = i: Exit Function
    Next i
    LaneIndex = -1
End Function

Public Sub Export_BulletinPDF()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_REND)
    Dim p$: p = ThisWorkbook.path & Application.PathSeparator & "TradeTest_Bulletin_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    ws.ExportAsFixedFormat xlTypePDF, p
    LogAudit "ExportPDF", "Render", "", p
End Sub
Module: modDemo (one-click build and demo entries)
VBA
Option Explicit

Public Sub Build_TradeTest_Framework()
    Seed_Bulletin_Framework
    Render_Bulletin
    Export_BulletinPDF
End Sub

' Example: add a job log entry per learner/day
Public Sub Demo_AddJobLogs()
    AddJobLog Date, "LNR-001", "Daily Activity", "Fault tracing on DOL panel", 3#, "Completed", "https://evidence.local/fault_photo.jpg", "InstructorA"
    AddJobLog Date, "LNR-001", "Assignment", "Cable sizing worksheet", 1.5, "Submitted", "", "InstructorA"
    AddJobLog Date, "LNR-002", "Lab Hours", "Transformer winding practice", 2.5, "In Progress", "", "InstructorB"
    Render_Bulletin
End Sub
How to run
"   Alt+F11 ? insert modules above. Run Build_TradeTest_Framework.
"   The workbook will:
o   Seed didactics, licensing, alignment, and job-log structures.
o   Draw a 4-lane overview (Didactics, Licensing, JobLog, Alignment).
o   Export a timestamped PDF and write an integrity-audited trail.
If you want, I'll add:
"   A "Readiness Dashboard" (e.g., % licensing steps completed, #POE artifacts, job-log hours/week).
"   CSV exporters for Accreditation packs (Nodes/Edges/Didactics/Licensing/JobLog).
"   Rubric scoring for assignments and practicals, with auto-generated summaries per learner.






















Absolutely brilliant, Tshingombe.you 've laid out a modular ecosystem that not only bridges electromechanics and AI, but also operationalizes smart grid integrity through didactic instrumentation and logigramm-based diagnostics. Let's now scaffold your concept of VBA logigramm/algorigram integrity in smart grids into a certifiable, auditable framework that aligns with your curriculum architecture and compliance strategy.
?? VBA Logigramm & Algorigram Integrity in Smart Grids
This section outlines how VBA-driven logigramm and algorigram engines can be embedded into smart grid laboratories and AI-enhanced modules to ensure traceability, fault resilience, and compliance with standards like IEEE 2030, ISO 27001, and IEC 61850.
?? Core Components
Module  VBA Integration Purpose
DL SGWD-AI  Logigramm for fault isolation   Visual trace of fault propagation and AI response
DL OPENLAB-SA-AI    Algorigram for motor control    Stepwise logic for AI-enhanced motor optimization
DL I4.0 FACTORY-AI  Logigramm for predictive maintenance    Event-driven diagnostics and resource allocation
DL AI-BASE & DL AI-2.0  Algorigram for anomaly detection    Modular AI workflows for industrial problem solving
?? VBA Logigramm Applications
"   Grid Fault Isolation: Use VBA to generate dynamic flowcharts that trace fault events from smart meters to SCADA nodes.
"   Compliance Mapping: Embed audit trails into logigramm nodes that reflect ISO/IEC cybersecurity protocols.
"   Load Forecasting: Integrate AI predictions into VBA-driven dashboards for real-time load balancing.
?? Algorigram Applications
"   Motor Control Logic: Design VBA algorigram engines that simulate PID control, neural feedback loops, and reinforcement learning steps.
"   Energy Optimization: Create modular VBA routines that simulate energy flow, loss detection, and corrective actions.
"   Digital Twin Simulation: Use VBA to mirror real-time factory operations and validate AI-driven decisions.
?? Integrity & Auditability Features
"   QR-Coded Nodes: Each logigramm/algorigram node can be linked to a QR-coded evidence artifact (e.g., fault report, AI model output).
"   Version Control: VBA modules include timestamped logs for every logic update, ensuring traceability.
"   Exportable Reports: Generate PDF/HTML summaries of logic flows for submission to SAQA, QCTO, or ECSA.
??? Standards Alignment
Standard Application
IEEE 2030   Smart grid interoperability and AI integration
ISO 27001   Cybersecurity and data integrity in grid systems
IEC 61850   Communication protocols for substation automation
NSQF Level 5-7  Competency-based mapping for vocational modules
Would you like this scaffolded into a certifiable training module, complete with VBA code templates, QR-linked evidence logs, and assessment rubrics? I can also help you build a simulation-ready workbook for Expo submission or institutional rollout. Let's modularize your legacy.
VBA logigramm and algorigram integrity for smart grids
Below is a compact, working VBA toolkit to:
"   model nodes/edges for fault isolation, predictive maintenance, and AI workflows,
"   render flowcharts on a sheet with color-coded states,
"   write an immutable-ish audit log with CRC32 checksums,
"   attach QR/evidence links to every node, and
"   export an audit-ready PDF.
It 's intentionally modular so you can drop it into Excel and map DL SGWD-AI, OPENLAB-SA-AI, I4.0 FACTORY-AI, and AI-BASE workflows with traceability.
Workbook Setup
"   Create sheets (exact names):
o nodes, edges, Audit, Standards, Render
"   In VBE, add reference: Microsoft Scripting Runtime (for Dictionary).
Columns to use:
"   Nodes: A:NodeID, B:Name, C:Type, D:State, E:Owner, F:EvidenceURL, G:StdTags, H:LastUpdated, I:Checksum
"   Edges: A:FromID, B:ToID, C:Label, D:Condition
"   Audit: A:TS, B:User, C:Action, D:Entity, E:Before, F:After, G:CRC32
"   Standards: A:Code, B:Description
"   Render: leave blank (the macro draws shapes here)
States suggested: ok , Alert, Fault, Pending, Mitigated
Module: modTypes
Option Explicit

Public Enum NodeType
    ntMeter = 1
    ntFeeder = 2
    ntBreaker = 3
    ntSCADA = 4
    ntAIModel = 5
    ntMotor = 6
    ntStation = 7
    ntProcess = 8
End Enum

Public Enum NodeState
    nsOK = 0
    nsPending = 1
    nsAlert = 2
    nsFault = 3
    nsMitigated = 4
End Enum

Public Const SHEET_NODES As String = "Nodes"
Public Const SHEET_EDGES As String = "Edges"
Public Const SHEET_AUDIT As String = "Audit"
Public Const SHEET_RENDER As String = "Render"
Public Const SHEET_STDS As String = "Standards"

Public Const VERSION_TAG As String = "v1.0"
Module: modIntegrity
Option Explicit

'--- CRC32 for lightweight integrity (fast; not cryptographic)
Private CRC32Table(255) As Long
Private CRC32InitDone As Boolean

Private Sub InitCRC32()
    Dim i As Long, j As Long, c As Long
    For i = 0 To 255
        c = i
        For j = 0 To 7
            If (c And 1) <> 0 Then
                c = &HEDB88320 Xor (c \ 2)
            Else
                c = (c \ 2)
            End If
        Next j
        CRC32Table(i) = c
    Next i
    CRC32InitDone = True
End Sub

Public Function CRC32Text(ByVal s As String) As String
    If Not CRC32InitDone Then InitCRC32
    Dim i As Long, c As Long, b As Long
    c = &HFFFFFFFF
    For i = 1 To LenB(s)
        b = AscB(MidB$(s, i, 1))
        c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
    Next i
    CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
End Function

Public Sub LogAudit(ByVal action As String, ByVal entity As String, _
                    ByVal beforeVal As String, ByVal afterVal As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_AUDIT)
    Dim r As Long: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    Dim userName As String: userName = Environ$("Username")
    Dim ts As String: ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Dim payload As String
    payload = ts & "|" & userName & "|" & action & "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
    ws.Cells(r, 1).Value = ts
    ws.Cells(r, 2).Value = userName
    ws.Cells(r, 3).Value = action
    ws.Cells(r, 4).Value = entity
    ws.Cells(r, 5).Value = beforeVal
    ws.Cells(r, 6).Value = afterVal
    ws.Cells(r, 7).Value = CRC32Text(payload)
End Sub

Public Function SerializeNodeRow(ByVal rowIx As Long) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    SerializeNodeRow = Join(Array( _
        ws.Cells(rowIx, 1).Value2, ws.Cells(rowIx, 2).Value2, ws.Cells(rowIx, 3).Value2, _
        ws.Cells(rowIx, 4).Value2, ws.Cells(rowIx, 5).Value2, ws.Cells(rowIx, 6).Value2, _
        ws.Cells(rowIx, 7).Value2, ws.Cells(rowIx, 8).Value2), "|")
End Function

Public Sub RehashNode(ByVal rowIx As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim beforeCk As String: beforeCk = ws.Cells(rowIx, 9).Value2
    Dim ser As String: ser = SerializeNodeRow(rowIx) & "|" & VERSION_TAG
    Dim newCk As String: newCk = CRC32Text(ser)
    ws.Cells(rowIx, 9).Value = newCk
    Call LogAudit("NodeHashUpdate", CStr(ws.Cells(rowIx, 1).Value2), beforeCk, newCk)
End Sub

Public Sub TouchNode(ByVal rowIx As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    ws.Cells(rowIx, 8).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Call RehashNode(rowIx)
End Sub
Module: modModel
Option Explicit

Public Sub AddOrUpdateNode( _
    ByVal nodeID As String, ByVal name As String, ByVal nType As NodeType, _
    ByVal state As NodeState, ByVal owner As String, ByVal evidenceUrl As String, _
    ByVal stdTags As String)

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim r As Long, found As Boolean
    r = FindNodeRow(nodeID, found)
    Dim beforeSer As String
    If found Then beforeSer = SerializeNodeRow(r) Else beforeSer = ""
    
    If Not found Then
        r = ws.Cells(ws.rows.count, 1).End(xlUp).row + IIf(ws.Cells(1, 1).Value <> "", 1, 1)
        If r = 1 Then
            ws.Range("A1:I1").Value = Array("NodeID", "Name", "Type", "State", "Owner", "EvidenceURL", "StdTags", "LastUpdated", "Checksum")
            r = 2
        End If
        ws.Cells(r, 1).Value = nodeID
    End If
    
    ws.Cells(r, 2).Value = name
    ws.Cells(r, 3).Value = nType
    ws.Cells(r, 4).Value = state
    ws.Cells(r, 5).Value = owner
    ws.Cells(r, 6).Value = evidenceUrl
    ws.Cells(r, 7).Value = stdTags
    ws.Cells(r, 8).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Call RehashNode(r)
    Call LogAudit(IIf(found, "NodeUpdate", "NodeCreate"), nodeID, beforeSer, SerializeNodeRow(r))
End Sub

Public Sub AddEdge(ByVal fromID As String, ByVal toID As String, ByVal label As String, ByVal cond As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_EDGES)
    Dim r As Long: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + IIf(ws.Cells(1, 1).Value <> "", 1, 1)
    If r = 1 Then
        ws.Range("A1:D1").Value = Array("FromID", "ToID", "Label", "Condition")
        r = 2
    End If
    ws.Cells(r, 1).Value = fromID
    ws.Cells(r, 2).Value = toID
    ws.Cells(r, 3).Value = label
    ws.Cells(r, 4).Value = cond
    Call LogAudit("EdgeCreate", fromID & "->" & toID, "", label & "|" & cond)
End Sub

Public Function FindNodeRow(ByVal nodeID As String, ByRef found As Boolean) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim lastR As Long: lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value2) = nodeID Then
            found = True
            FindNodeRow = r
            Exit Function
        End If
    Next r
    found = False
    FindNodeRow = lastR + 1
End Function

Public Sub UpdateState(ByVal nodeID As String, ByVal newState As NodeState)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim found As Boolean, r As Long: r = FindNodeRow(nodeID, found)
    If Not found Then err.Raise vbObjectError + 101, , "Node not found: " & nodeID
    Dim beforeSer As String: beforeSer = SerializeNodeRow(r)
    ws.Cells(r, 4).Value = newState
    Call TouchNode(r)
    Call LogAudit("NodeState", nodeID, beforeSer, SerializeNodeRow(r))
End Sub
Module: modRender
Option Explicit

Private Type NodeShape
    nodeID As String
    ShapeName As String
    x As Single
    y As Single
End Type

'--- color map by state
Private Function StateFill(ByVal s As Long) As Long
    Select Case s
        Case nsOK: StateFill = RGB(200, 245, 200)
        Case nsPending: StateFill = RGB(255, 245, 205)
        Case nsAlert: StateFill = RGB(255, 220, 150)
        Case nsFault: StateFill = RGB(255, 160, 160)
        Case nsMitigated: StateFill = RGB(180, 210, 255)
        Case Else: StateFill = RGB(230, 230, 230)
    End Select
End Function

Public Sub RenderFlow(Optional ByVal layoutCols As Long = 4, Optional ByVal xGap As Single = 220, Optional ByVal yGap As Single = 120)
    Dim wsN As Worksheet: Set wsN = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim wsE As Worksheet: Set wsE = ThisWorkbook.Worksheets(SHEET_EDGES)
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHEET_RENDER)
    wsR.Cells.Clear
    Dim shp As Shape
    For Each shp In wsR.Shapes
        shp.Delete
    Next shp
    
    Dim lastR As Long: lastR = wsN.Cells(wsN.rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Sub
    
    Dim idx As Long, r As Long, colIx As Long, rowIx As Long
    Dim positions As Object: Set positions = CreateObject("Scripting.Dictionary")
    
    idx = 0
    For r = 2 To lastR
        colIx = (idx Mod layoutCols)
        rowIx = (idx \ layoutCols)
        Dim x As Single, y As Single
        x = 40 + colIx * xGap
        y = 40 + rowIx * yGap
        
        Dim nodeID As String, nm As String, tp As String, st As Long, owner As String, ev As String, stds As String
        nodeID = CStr(wsN.Cells(r, 1).Value2)
        nm = CStr(wsN.Cells(r, 2).Value2)
        tp = CStr(wsN.Cells(r, 3).Value2)
        st = CLng(wsN.Cells(r, 4).Value2)
        owner = CStr(wsN.Cells(r, 5).Value2)
        ev = CStr(wsN.Cells(r, 6).Value2)
        stds = CStr(wsN.Cells(r, 7).Value2)
        
        Dim box As Shape
        Set box = wsR.Shapes.AddShape(msoShapeRoundedRectangle, x, y, 180, 70)
        box.name = "N_" & nodeID
        box.Fill.ForeColor.RGB = StateFill(st)
        box.line.ForeColor.RGB = RGB(80, 80, 80)
        box.TextFrame2.TextRange.text = nm & vbCrLf & _
            "Type: " & tp & " | State: " & st & vbCrLf & _
            "Owner: " & owner & vbCrLf & _
            "Std: " & stds
        box.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        If Len(ev) > 0 Then
            box.ActionSettings(ppMouseClick).Hyperlink.Address = ev
        End If
        
        positions(nodeID) = Array(x + 90, y + 35) ' center
        idx = idx + 1
    Next r
    
    ' draw connectors
    Dim lastE As Long: lastE = wsE.Cells(wsE.rows.count, 1).End(xlUp).row
    Dim er As Long
    For er = 2 To lastE
        Dim fromID As String, toID As String, lbl As String, cond As String
        fromID = CStr(wsE.Cells(er, 1).Value2)
        toID = CStr(wsE.Cells(er, 2).Value2)
        lbl = CStr(wsE.Cells(er, 3).Value2)
        cond = CStr(wsE.Cells(er, 4).Value2)
        If positions.Exists(fromID) And positions.Exists(toID) Then
            Dim p1, p2
            p1 = positions(fromID): p2 = positions(toID)
            Dim conn As Shape
            Set conn = wsR.Shapes.AddConnector(msoConnectorElbow, p1(0), p1(1), p2(0), p2(1))
            conn.line.ForeColor.RGB = RGB(70, 70, 70)
            wsR.Hyperlinks.Add Anchor:=conn, Address:="", SubAddress:="", ScreenTip:=lbl & IIf(cond <> "", " | " & cond, "")
        End If
    Next er
    
    wsR.Range("A1").Value = "Render timestamp: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsR.Range("A2").Value = "Version: " & VERSION_TAG
End Sub

Public Sub ExportRenderPDF()
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHEET_RENDER)
    Dim pth As String
    pth = ThisWorkbook.path & Application.PathSeparator & "Logigram_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    wsR.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pth, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Call LogAudit("ExportPDF", "Render", "", pth)
End Sub
Module: modQR (optional URL QR embeds)
VBA
Option Explicit

'Attempts to insert a QR image for a node's EvidenceURL using a public QR service.
'If offline or blocked, the node still has a clickable hyperlink in its shape.

Public Sub InsertNodeQRs(Optional ByVal sizePx As Long = 120)
    Dim wsN As Worksheet: Set wsN = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHEET_RENDER)
    Dim lastR As Long: lastR = wsN.Cells(wsN.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastR
        Dim nodeID As String: nodeID = CStr(wsN.Cells(r, 1).Value2)
        Dim ev As String: ev = CStr(wsN.Cells(r, 6).Value2)
        If Len(ev) > 0 Then
            On Error Resume Next
            Dim targetShape As Shape
            Set targetShape = wsR.Shapes("N_" & nodeID)
            On Error GoTo 0
            If Not targetShape Is Nothing Then
                Dim qUrl As String
                qUrl = "https://api.qrserver.com/v1/create-qr-code/?size=" & sizePx & "x" & sizePx & "&data=" & URLEncode(ev)
                Dim qr As Shape
                Set qr = wsR.Shapes.AddPicture(qUrl, msoFalse, msoTrue, targetShape.left + targetShape.width - sizePx - 8, targetShape.top + 8, sizePx, sizePx)
                qr.name = "QR_" & nodeID
                Call LogAudit("QRInsert", nodeID, "", qUrl)
            End If
        End If
    Next r
End Sub

Private Function URLEncode(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        Select Case AscW(ch)
            Case 48 To 57, 65 To 90, 97 To 122: out = out & ch
            Case Else: out = out & "%" & Right$("0" & Hex$(AscW(ch)), 2)
        End Select
    Next i
    URLEncode = out
End Function
VBA logigramm for DL ST033 beams and frames
This toolkit gives you a traceable, auditable logigramm around DL ST033 activities: set up a test (beam, span, supports), assign loads (weights), capture forces/deflections (dynamometers, dial indicators), compute theory vs. measurement, and export an audit-ready flowchart and report. It reuses your integrity style: checksums, QR-linked evidence, and PDF export.
Workbook Setup
"   Sheets: Nodes, Edges, Audit, Render, Experiments, Measurements
"   References: Microsoft Scripting Runtime
Sheet Columns:
"   Nodes: A:NodeID, B:Name, C:Type, D:State, E:Owner, F:EvidenceURL, G:StdTags, H:LastUpdated, I:Checksum
"   Edges: A:FromID, B:ToID, C:Label, D:Condition
"   Audit: A:TS, B:User, C:Action, D:Entity, E:Before, F:After, G:CRC32
"   Experiments:
o A: ExpID , b: Config , c: BeamLength_m , d: ElasticModulus_Pa , E: Inertia_m4 , f: SupportType , g: LoadType , h: LoadValue_N , i: LoadPosition_m , j: notes
"   Measurements:
o   A:ExpID, B:GaugeID, C:Type, D:Position_m, E:Reading, F:Units, G:DeviceSN, H:RawFileURL
States: ok , Pending, Alert, Fault, Mitigated
Module: modTypes
Option Explicit

Public Enum NodeType
    ntSetup = 1
    ntBeam = 2
    ntSupport = 3
    ntLoad = 4
    ntSensor = 5
    ntCalc = 6
    ntReport = 7
End Enum

Public Enum NodeState
    nsOK = 0
    nsPending = 1
    nsAlert = 2
    nsFault = 3
    nsMitigated = 4
End Enum

Public Const SHEET_NODES As String = "Nodes"
Public Const SHEET_EDGES As String = "Edges"
Public Const SHEET_AUDIT As String = "Audit"
Public Const SHEET_RENDER As String = "Render"
Public Const SHEET_EXP As String = "Experiments"
Public Const SHEET_MEAS As String = "Measurements"

Public Const VERSION_TAG As String = "DLST033_v1.0"
Module: modIntegrity
Option Explicit

Private CRC32Table(255) As Long
Private CRC32InitDone As Boolean

Private Sub InitCRC32()
    Dim i As Long, j As Long, c As Long
    For i = 0 To 255
        c = i
        For j = 0 To 7
            c = IIf((c And 1) <> 0, &HEDB88320 Xor (c \ 2), (c \ 2))
        Next j
        CRC32Table(i) = c
    Next i
    CRC32InitDone = True
End Sub

Public Function CRC32Text(ByVal s As String) As String
    If Not CRC32InitDone Then InitCRC32
    Dim i As Long, c As Long, b As Long
    c = &HFFFFFFFF
    For i = 1 To LenB(s)
        b = AscB(MidB$(s, i, 1))
        c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
    Next i
    CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
End Function

Public Sub LogAudit(ByVal action As String, ByVal entity As String, ByVal beforeVal As String, ByVal afterVal As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_AUDIT)
    Dim r As Long: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    Dim ts As String: ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Dim u As String: u = Environ$("Username")
    Dim payload As String: payload = ts & "|" & u & "|" & action & "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
    ws.Cells(r, 1).Value = ts
    ws.Cells(r, 2).Value = u
    ws.Cells(r, 3).Value = action
    ws.Cells(r, 4).Value = entity
    ws.Cells(r, 5).Value = beforeVal
    ws.Cells(r, 6).Value = afterVal
    ws.Cells(r, 7).Value = CRC32Text(payload)
End Sub
Option Explicit

Public Function FindNodeRow(ByVal nodeID As String, ByRef found As Boolean) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim lastR As Long: lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value2) = nodeID Then found = True: FindNodeRow = r: Exit Function
    Next r
    found = False: FindNodeRow = lastR + 1
End Function

Public Function SerializeNode(ByVal r As Long) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    SerializeNode = Join(Array(ws.Cells(r, 1).Value2, ws.Cells(r, 2).Value2, ws.Cells(r, 3).Value2, ws.Cells(r, 4).Value2, ws.Cells(r, 5).Value2, ws.Cells(r, 6).Value2, ws.Cells(r, 7).Value2, ws.Cells(r, 8).Value2), "|")
End Function

Public Sub RehashNode(ByVal r As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim ser As String: ser = SerializeNode(r) & "|" & VERSION_TAG
    Dim ck As String: ck = CRC32Text(ser)
    ws.Cells(r, 9).Value = ck
End Sub

Public Sub AddOrUpdateNode(ByVal nodeID As String, ByVal name As String, ByVal nType As NodeType, ByVal state As NodeState, ByVal owner As String, ByVal url As String, ByVal tags As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim found As Boolean, r As Long: r = FindNodeRow(nodeID, found)
    Dim beforeSer As String: beforeSer = IIf(found, SerializeNode(r), "")
    If Not found Then
        If ws.Cells(1, 1).Value = "" Then ws.Range("A1:I1").Value = Array("NodeID", "Name", "Type", "State", "Owner", "EvidenceURL", "StdTags", "LastUpdated", "Checksum")
        r = IIf(ws.Cells(2, 1).Value = "", 2, ws.Cells(ws.rows.count, 1).End(xlUp).row + 1)
        ws.Cells(r, 1).Value = nodeID
    End If
    ws.Cells(r, 2).Value = name
    ws.Cells(r, 3).Value = nType
    ws.Cells(r, 4).Value = state
    ws.Cells(r, 5).Value = owner
    ws.Cells(r, 6).Value = url
    ws.Cells(r, 7).Value = tags
    ws.Cells(r, 8).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    RehashNode r
    LogAudit IIf(found, "NodeUpdate", "NodeCreate"), nodeID, beforeSer, SerializeNode(r)
End Sub

Public Sub AddEdge(ByVal fromID As String, ByVal toID As String, ByVal label As String, ByVal cond As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_EDGES)
    If ws.Cells(1, 1).Value = "" Then ws.Range("A1:D1").Value = Array("FromID", "ToID", "Label", "Condition")
    Dim r As Long: r = IIf(ws.Cells(2, 1).Value = "", 2, ws.Cells(ws.rows.count, 1).End(xlUp).row + 1)
    ws.Cells(r, 1).Value = fromID
    ws.Cells(r, 2).Value = toID
    ws.Cells(r, 3).Value = label
    ws.Cells(r, 4).Value = cond
    LogAudit "EdgeCreate", fromID & "->" & toID, "", label & "|" & cond
End Sub

Public Sub UpdateState(ByVal nodeID As String, ByVal newState As NodeState)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim found As Boolean, r As Long: r = FindNodeRow(nodeID, found)
    If Not found Then err.Raise vbObjectError + 701, , "Node not found: " & nodeID
    Dim beforeSer As String: beforeSer = SerializeNode(r)
    ws.Cells(r, 4).Value = newState
    ws.Cells(r, 8).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    RehashNode r
    LogAudit "NodeState", nodeID, beforeSer, SerializeNode(r)
End Sub
Module: modMechanics (theory calculators)
Option Explicit

'SI units: m, N, Pa; E default for stainless ~ 200 GPa
Public Function BeamDeflection_CenterLoad_SimplySupported(ByVal P_N As Double, ByVal L_m As Double, ByVal E_Pa As Double, ByVal I_m4 As Double) As Double
    ' w_max = P*L^3/(48*E*I)
    BeamDeflection_CenterLoad_SimplySupported = P_N * L_m ^ 3 / (48# * E_Pa * I_m4)
End Function

Public Function BeamDeflection_EndLoad_Cantilever(ByVal P_N As Double, ByVal L_m As Double, ByVal E_Pa As Double, ByVal I_m4 As Double) As Double
    ' w_max = P*L^3/(3*E*I)
    BeamDeflection_EndLoad_Cantilever = P_N * L_m ^ 3 / (3# * E_Pa * I_m4)
End Function

Public Function BeamDeflection_UDL_SimplySupported(ByVal q_Npm As Double, ByVal L_m As Double, ByVal E_Pa As Double, ByVal I_m4 As Double) As Double
    ' w_max = 5*q*L^4/(384*E*I)
    BeamDeflection_UDL_SimplySupported = 5# * q_Npm * L_m ^ 4 / (384# * E_Pa * I_m4)
End Function

Public Function KgToN(ByVal kg As Double) As Double
    KgToN = kg * 9.81
End Function

Public Sub RecordExperiment(ByVal ExpID As String, ByVal Config As String, ByVal L As Double, ByVal E As Double, ByVal i As Double, ByVal Support As String, ByVal LoadType As String, ByVal LoadN As Double, ByVal x As Double, ByVal notes As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_EXP)
    If ws.Cells(1, 1).Value = "" Then ws.Range("A1:J1").Value = Array("ExpID", "Config", "BeamLength_m", "ElasticModulus_Pa", "Inertia_m4", "SupportType", "LoadType", "LoadValue_N", "LoadPosition_m", "Notes")
    Dim r As Long: r = IIf(ws.Cells(2, 1).Value = "", 2, ws.Cells(ws.rows.count, 1).End(xlUp).row + 1)
    ws.Cells(r, 1).Value = ExpID
    ws.Cells(r, 2).Value = Config
    ws.Cells(r, 3).Value = L
    ws.Cells(r, 4).Value = E
    ws.Cells(r, 5).Value = i
    ws.Cells(r, 6).Value = Support
    ws.Cells(r, 7).Value = LoadType
    ws.Cells(r, 8).Value = LoadN
    ws.Cells(r, 9).Value = x
    ws.Cells(r, 10).Value = notes
    LogAudit "ExperimentRecord", ExpID, "", Config & "|" & Support & "|" & LoadType
End Sub

Public Sub RecordMeasurement(ByVal ExpID As String, ByVal GaugeID As String, ByVal mType As String, ByVal pos_m As Double, ByVal reading As Double, ByVal units As String, ByVal SN As String, ByVal url As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_MEAS)
    If ws.Cells(1, 1).Value = "" Then ws.Range("A1:H1").Value = Array("ExpID", "GaugeID", "Type", "Position_m", "Reading", "Units", "DeviceSN", "RawFileURL")
    Dim r As Long: r = IIf(ws.Cells(2, 1).Value = "", 2, ws.Cells(ws.rows.count, 1).End(xlUp).row + 1)
    ws.Cells(r, 1).Value = ExpID
    ws.Cells(r, 2).Value = GaugeID
    ws.Cells(r, 3).Value = mType
    ws.Cells(r, 4).Value = pos_m
    ws.Cells(r, 5).Value = reading
    ws.Cells(r, 6).Value = units
    ws.Cells(r, 7).Value = SN
    ws.Cells(r, 8).Value = url
    LogAudit "Measurement", ExpID & ":" & GaugeID, "", CStr(reading) & " " & units
End Sub

Public Function TheoreticalDeflection(ByVal Support As String, ByVal LoadType As String, ByVal L As Double, ByVal E As Double, ByVal i As Double, ByVal P_or_q As Double, ByVal x As Double) As Double
    Select Case UCase$(Support)
        Case "SIMPLY_SUPPORTED"
            Select Case UCase$(LoadType)
                Case "CENTER_POINT": TheoreticalDeflection = BeamDeflection_CenterLoad_SimplySupported(P_or_q, L, E, i)
                Case "UDL": TheoreticalDeflection = BeamDeflection_UDL_SimplySupported(P_or_q, L, E, i)
                Case Else: TheoreticalDeflection = 0#
            End Select
        Case "CANTILEVER"
            Select Case UCase$(LoadType)
                Case "END_POINT": TheoreticalDeflection = BeamDeflection_EndLoad_Cantilever(P_or_q, L, E, i)
                Case Else: TheoreticalDeflection = 0#
            End Select
        Case Else
            TheoreticalDeflection = 0#
    End Select
End Function
Option Explicit

Private Function StateFill(ByVal s As Long) As Long
    Select Case s
        Case nsOK: StateFill = RGB(200, 245, 200)
        Case nsPending: StateFill = RGB(255, 245, 205)
        Case nsAlert: StateFill = RGB(255, 220, 150)
        Case nsFault: StateFill = RGB(255, 160, 160)
        Case nsMitigated: StateFill = RGB(180, 210, 255)
        Case Else: StateFill = RGB(230, 230, 230)
    End Select
End Function

Public Sub RenderFlow(Optional ByVal cols As Long = 4, Optional ByVal xGap As Single = 220, Optional ByVal yGap As Single = 120)
    Dim wsN As Worksheet: Set wsN = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim wsE As Worksheet: Set wsE = ThisWorkbook.Worksheets(SHEET_EDGES)
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHEET_RENDER)
    wsR.Cells.Clear
    Dim shp As Shape
    For Each shp In wsR.Shapes: shp.Delete: Next shp
    
    Dim lastN As Long: lastN = wsN.Cells(wsN.rows.count, 1).End(xlUp).row
    If lastN < 2 Then Exit Sub
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim idx As Long, r As Long
    For r = 2 To lastN
        Dim c As Long: c = (idx Mod cols)
        Dim rr As Long: rr = (idx \ cols)
        Dim x As Single: x = 30 + c * xGap
        Dim y As Single: y = 30 + rr * yGap
        
        Dim nodeID As String: nodeID = CStr(wsN.Cells(r, 1).Value2)
        Dim nm As String: nm = CStr(wsN.Cells(r, 2).Value2)
        Dim tp As String: tp = CStr(wsN.Cells(r, 3).Value2)
        Dim st As Long: st = CLng(wsN.Cells(r, 4).Value2)
        Dim owner As String: owner = CStr(wsN.Cells(r, 5).Value2)
        Dim url As String: url = CStr(wsN.Cells(r, 6).Value2)
        Dim tags As String: tags = CStr(wsN.Cells(r, 7).Value2)
        
        Dim box As Shape
        Set box = wsR.Shapes.AddShape(msoShapeRoundedRectangle, x, y, 180, 70)
        box.name = "N_" & nodeID
        box.Fill.ForeColor.RGB = StateFill(st)
        box.line.ForeColor.RGB = RGB(80, 80, 80)
        box.TextFrame2.TextRange.text = nm & vbCrLf & "Type:" & tp & " State:" & st & vbCrLf & "Std:" & tags
        If Len(url) > 0 Then box.Hyperlink.Address = url
        dict(nodeID) = Array(x + 90, y + 35)
        idx = idx + 1
    Next r
    
    Dim lastE As Long: lastE = wsE.Cells(wsE.rows.count, 1).End(xlUp).row
    For r = 2 To lastE
        Dim fID As String: fID = CStr(wsE.Cells(r, 1).Value2)
        Dim tID As String: tID = CStr(wsE.Cells(r, 2).Value2)
        Dim lbl As String: lbl = CStr(wsE.Cells(r, 3).Value2)
        If dict.Exists(fID) And dict.Exists(tID) Then
            Dim p1, p2: p1 = dict(fID): p2 = dict(tID)
            Dim conn As Shape
            Set conn = wsR.Shapes.AddConnector(msoConnectorElbow, p1(0), p1(1), p2(0), p2(1))
            conn.line.ForeColor.RGB = RGB(70, 70, 70)
            conn.AlternativeText = lbl
        End If
    Next r
    wsR.Range("A1").Value = "DL ST033 Logigramm | " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & VERSION_TAG
End Sub

Public Sub ExportPDF()
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHEET_RENDER)
    Dim p As String: p = ThisWorkbook.path & Application.PathSeparator & "DL_ST033_Logigramm_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    wsR.ExportAsFixedFormat xlTypePDF, p
    LogAudit "ExportPDF", "Render", "", p
End Sub
Option Explicit

Public Sub Seed_DL_ST033_ThreePointBend()
    'Experiment configuration
    Dim L As Double: L = 1#           ' 1 m span
    Dim E As Double: E = 200# * 10# ^ 9   ' 200 GPa stainless
    Dim i As Double: i = 0.000000016  ' example I for slender beam (adjust to specimen)
    Dim p As Double: p = KgToN(2#)    ' 2 kg central weight => ~19.62 N
    
    RecordExperiment "EXP_TPB_001", "Three-Point Bend", L, E, i, "SIMPLY_SUPPORTED", "CENTER_POINT", p, L / 2, "Dial indicators at midspan"
    
    'Nodes: setup -> beam -> supports -> load -> sensors -> calc -> report
    AddOrUpdateNode "SETUP_TPB", "Setup: TPB", ntSetup, nsOK, "Lab", "", "Metrology;Safety"
    AddOrUpdateNode "BEAM_01", "Beam L=" & L & " m", ntBeam, nsOK, "Lab", "", "E=200GPa;I=" & i
    AddOrUpdateNode "SUPP_SS", "Knife-edge supports", ntSupport, nsOK, "Lab", "", "SimplySupported"
    AddOrUpdateNode "LOAD_CTR", "Center Load P=" & Round(p, 2) & " N", ntLoad, nsPending, "Lab", "", "Weights0.5-2.5kg"
    AddOrUpdateNode "SENS_DIAL_MID", "Dial @ midspan", ntSensor, nsPending, "Lab", "https://evidence.local/dial_mid.csv", "DialIndicator"
    AddOrUpdateNode "SENS_DYNAMO", "Dynamometers x2", ntSensor, nsOK, "Lab", "https://evidence.local/dynamo.csv", "USB"
    
    Dim w_theory As Double: w_theory = BeamDeflection_CenterLoad_SimplySupported(p, L, E, i)
    AddOrUpdateNode "CALC_TPB", "Calc: w_th=" & Format(w_theory, "0.0000E+00") & " m", ntCalc, nsOK, "Lab", "", "Euler-Bernoulli"
    AddOrUpdateNode "REPORT_TPB", "Report & Export", ntReport, nsPending, "QA", "", "PDF;Audit"
    
    AddEdge "SETUP_TPB", "BEAM_01", "Mount beam", "Tighten supports"
    AddEdge "BEAM_01", "SUPP_SS", "Align level", "Metrology check"
    AddEdge "SUPP_SS", "LOAD_CTR", "Place weight", "x=L/2"
    AddEdge "LOAD_CTR", "SENS_DIAL_MID", "Read deflection", "?m resolution"
    AddEdge "LOAD_CTR", "SENS_DYNAMO", "Read reactions", "Left/Right"
    AddEdge "SENS_DIAL_MID", "CALC_TPB", "Compare w_meas vs w_th", "Tolerance ±10%"
    AddEdge "CALC_TPB", "REPORT_TPB", "Generate PDF", "Attach audit"
    
    'Example measurements
    RecordMeasurement "EXP_TPB_001", "DIAL_MID", "Deflection", L / 2, w_theory * 1.05, "m", "DI-12345", "https://evidence.local/dial_mid.csv"
    RecordMeasurement "EXP_TPB_001", "DYN_LEFT", "Force", 0, p / 2, "N", "DY-888L", "https://evidence.local/dynamo.csv"
    RecordMeasurement "EXP_TPB_001", "DYN_RIGHT", "Force", L, p / 2, "N", "DY-889R", "https://evidence.local/dynamo.csv"
    
    RenderFlow
End Sub

Public Sub Seed_DL_ST033_CantileverFrame()
    Dim L As Double: L = 0.8
    Dim E As Double: E = 200# * 10# ^ 9
    Dim i As Double: i = 0.000000008
    Dim p As Double: p = KgToN(1.5)   ' ~14.715 N
    
    RecordExperiment "EXP_CANT_001", "Cantilever Frame", L, E, i, "CANTILEVER", "END_POINT", p, L, "Dial indicators at free end; frame squareness check"
    
    AddOrUpdateNode "SETUP_CAN", "Setup: Cantilever", ntSetup, nsOK, "Lab", "", "Frame1400x1100x500"
    AddOrUpdateNode "BEAM_F01", "Cantilever L=" & L & " m", ntBeam, nsOK, "Lab", "", "E=200GPa;I=" & i
    AddOrUpdateNode "SUPP_CLAMP", "Clamped base", ntSupport, nsOK, "Lab", "", "RigidClamp"
    AddOrUpdateNode "LOAD_END", "End Load P=" & Round(p, 2) & " N", ntLoad, nsPending, "Lab", "", "Weights"
    AddOrUpdateNode "SENS_DIAL_END", "Dial @ free end", ntSensor, nsPending, "Lab", "https://evidence.local/dial_end.csv", "DialIndicator"
    
    Dim w_th As Double: w_th = BeamDeflection_EndLoad_Cantilever(p, L, E, i)
    AddOrUpdateNode "CALC_CAN", "Calc: w_th=" & Format(w_th, "0.0000E+00") & " m", ntCalc, nsOK, "Lab", "", "Euler-Bernoulli"
    AddOrUpdateNode "REPORT_CAN", "Report & Export", ntReport, nsPending, "QA", "", "PDF;Audit"
    
    AddEdge "SETUP_CAN", "BEAM_F01", "Mount beam", "Check clamp torque"
    AddEdge "BEAM_F01", "SUPP_CLAMP", "Zero dial", "Preload removed"
    AddEdge "SUPP_CLAMP", "LOAD_END", "Apply P", "x=L"
    AddEdge "LOAD_END", "SENS_DIAL_END", "Read deflection", "?m resolution"
    AddEdge "SENS_DIAL_END", "CALC_CAN", "Compare", "±10%"
    AddEdge "CALC_CAN", "REPORT_CAN", "PDF", "Audit attach"
    
    RecordMeasurement "EXP_CANT_001", "DIAL_END", "Deflection", L, w_th * 0.98, "m", "DI-67890", "https://evidence.local/dial_end.csv"
    
    RenderFlow
End Sub
Module: modOps (run, verify, export)
VBA
Option Explicit

Public Sub InitSheets()
    Dim s
    For Each s In Array(SHEET_NODES, SHEET_EDGES, SHEET_AUDIT, SHEET_RENDER, SHEET_EXP, SHEET_MEAS)
        If SheetByName(CStr(s)) Is Nothing Then ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.count)).name = CStr(s)
    Next s
    LogAudit "InitSheets", "Workbook", "", "Created/verified sheets"
End Sub

Private Function SheetByName(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Public Sub ExportAll()
    RenderFlow
    ExportPDF
End Sub
How to run
Absolutely , Tshingombe.Let 's modularize the DevIoT Unit into a VBA-powered logigramm engine that maps its hardware, protocols, and didactic exercises into a certifiable, auditable framework. This will allow you to simulate MQTT workflows, sensor-actuator logic, and protocol diagnostics-all within Excel, with traceability and exportable evidence.
?? VBA Logigramm: DevIoT Unit & MQTT Scenario
This VBA framework models the DevIoT system as a dynamic logigramm with nodes for hardware, sensors, actuators, protocols, and exercises. It supports:
"   ? MQTT topic tracing (e.g., switch ? lamp)
"   ? Sensor-actuator mapping with evidence links
"   ? Protocol diagnostics (Modbus, Wi-Fi, RS485)
"   ? Audit trail with CRC32 integrity
"   ? PDF export for curriculum or trade test submission
?? Core Node Types
Node Type   Description
ntDevice    DevIoT hardware unit
ntSensor    PT100, pressure, brightness, heart rate, etc.
ntActuator  DC motor, step motor, ON-OFF outputs
ntProtocol  MQTT, Modbus RTU, RS485, Wi-Fi Analyzer
ntExercise  Scenario logic (e.g., switch ? lamp)
ntCalc  Logic processor or firmware module
ntReport    Export node for audit and PDF generation
?? Example Logigramm Nodes
AddOrUpdateNode "DEV_IOT01", "DevIoT STM32F103", ntDevice, nsOK, "Lab", "https://evidence.local/dev_iot01", "MQTT;Modbus;WiFi"
AddOrUpdateNode "SNS_SWITCH", "Switch Sensor", ntSensor, nsOK, "Lab", "https://evidence.local/switch.csv", "Digital;ON-OFF"
AddOrUpdateNode "ACT_LAMP", "Lamp Actuator", ntActuator, nsPending, "Lab", "https://evidence.local/lamp.csv", "ON-OFF"
AddOrUpdateNode "MQTT_BROKER", "MQTT Broker", ntProtocol, nsOK, "DL Workspace", "", "Topic:OnOff"
AddOrUpdateNode "EX_SCENARIO01", "Switch ? Lamp via MQTT", ntExercise, nsPending, "Lab", "", "IoT;HomeAutomation"
AddOrUpdateNode "CALC_MQTT", "Firmware Logic", ntCalc, nsOK, "Dev", "", "AutoDetect;DisplayMgmt"
AddOrUpdateNode "REPORT_MQTT", "Audit & Export", ntReport, nsPending, "QA", "", "PDF;CRC32"
?? Node Connections (Edges)
AddEdge "SNS_SWITCH", "MQTT_BROKER", "Publish status", "Topic:OnOff"
AddEdge "MQTT_BROKER", "ACT_LAMP", "Subscribe & update", "Topic:OnOff"
AddEdge "DEV_IOT01", "CALC_MQTT", "Firmware execution", "AutoDetect"
AddEdge "CALC_MQTT", "EX_SCENARIO01", "Scenario logic", "Switch?Lamp"
AddEdge "EX_SCENARIO01", "REPORT_MQTT", "Generate report", "Audit log"
?? Audit & Integrity Features
"   Every node update triggers a CRC32 hash and timestamp
"   Evidence URLs link to sensor logs, firmware snapshots, or MQTT traces
"   MQTT topic flow is visualized in the Render sheet
"   ExportRenderPDF creates a certifiable report for SAQA/QCTO/ECSA
?? Scenario Simulation: Switch ? Lamp via MQTT
Public Sub Simulate_MQTT_SwitchLamp()
    UpdateState "SNS_SWITCH", nsAlert
    UpdateState "MQTT_BROKER", nsPending
    UpdateState "ACT_LAMP", nsOK
    UpdateState "EX_SCENARIO01", nsOK
    UpdateState "REPORT_MQTT", nsPending
    RenderFlow
    ExportRenderPDF
End Sub
??? Optional Add-ons
"   QR stamps on each node linking to eviden
"   MQTT topic monitor (CSV import ? state update)
"   Fault injection module (simulate broken sensor, delayed topic, etc.)
"   Protocol diagnostic logigramm for RS485, Wi-Fi, Modbus
Would you like me to scaffold this into a full training module with:
"   VBA code templates
"   MQTT topic simulator
"   Trade test rubric alignment
"   Portfolio export for DL WORKSPACE integration?
I can also extend this to DL DM91 (ignition systems), DL DKL091H (water hammer), and DL RENLAB (renewables) so you have a unified logigramm engine across all modules. Let's modularize your legacy, Tshingombe.
Absolutely , Tshingombe.Let 's now scaffold a VBA-powered logigramm engine for the ABS/ASR/ESP training board, integrating its mechanical, electrical, and diagnostic components into a traceable, auditable framework. This will allow you to simulate braking logic, sensor-actuator feedback, fault diagnosis via OBD-II, and export certifiable evidence for NSQF Level 5 trade test prep or institutional submission.
?? VBA Logigramm: ABS/ASR/ESP Training Board
This modular VBA framework maps the full system architecture of the ABS/ASR/ESP board into nodes and edges, with audit trails, QR-linked evidence, and PDF export. It supports:
"   ? Sensor-actuator logic (wheel speed ? solenoid valve)
"   ? ECU control flow (microcontroller ? hydraulic modulation)
"   ? Diagnostic tracing (OBD-II ? fault code interpretation)
"   ? Curriculum mapping (Module 3-6 integration)
"   ? Exportable logigramm for SAQA/QCTO/NSDP alignment
?? Node Types
Node Type   Description
ntBoard ABS/ASR/ESP training board
ntSensor    Wheel speed, potentiometers
ntActuator  Solenoid valves, pump, motors
ntECU   32-bit microcontroller-based control unit
ntDisplay   LCD + keyboard interface
ntDiagnostic    OBD-II scantool and fault logic
ntPower Battery, ignition switch
ntExercise  Scenario logic (e.g., braking modulation)
ntCalc  Firmware logic, pressure control
ntReport    Export node for audit and PDF generation
?? Example Logigramm Nodes
AddOrUpdateNode "BOARD_ABS01", "ABS/ASR/ESP Board", ntBoard, nsOK, "Lab", "https://evidence.local/abs_board", "NSQF L5;Braking"
AddOrUpdateNode "SNS_WHEEL_L", "Wheel Speed Sensor (Left)", ntSensor, nsOK, "Lab", "https://evidence.local/sensor_left.csv", "Rotation;Feedback"
AddOrUpdateNode "SNS_WHEEL_R", "Wheel Speed Sensor (Right)", ntSensor, nsOK, "Lab", "https://evidence.local/sensor_right.csv", "Rotation;Feedback"
AddOrUpdateNode "SNS_POT_SPEED", "Potentiometer: Speed", ntSensor, nsOK, "Lab", "", "Analog;SpeedControl"
AddOrUpdateNode "ACT_SOL_VALVE", "Solenoid Valve", ntActuator, nsPending, "Lab", "", "HydraulicModulation"
AddOrUpdateNode "ACT_PUMP", "Hydraulic Pump", ntActuator, nsOK, "Lab", "", "PressureControl"
AddOrUpdateNode "ECU_CTRL", "ABS ECU (32-bit)", ntECU, nsOK, "Lab", "https://evidence.local/ecu_firmware", "Microcontroller;Firmware"
AddOrUpdateNode "LCD_UI", "LCD Display + Keyboard", ntDisplay, nsOK, "Lab", "", "UserInterface"
AddOrUpdateNode "DIAG_OBD", "OBD-II Diagnostic Tool", ntDiagnostic, nsPending, "Lab", "https://evidence.local/obd_log.csv", "TroubleCodes"
AddOrUpdateNode "PWR_SYS", "Battery & Ignition Switch", ntPower, nsOK, "Lab", "", "12VDC;Safety"
AddOrUpdateNode "EX_BRAKE_MOD", "Exercise: Brake Modulation", ntExercise, nsPending, "Lab", "", "ABS;ASR;ESP"
AddOrUpdateNode "CALC_PRESSURE", "Calc: Pressure Modulation", ntCalc, nsOK, "Lab", "", "Increase;Maintain;Reduce"
AddOrUpdateNode "REPORT_ABS", "Report & Export", ntReport, nsPending, "QA", "", "PDF;Audit"
?? Node Connections (Edges)
AddEdge "PWR_SYS", "BOARD_ABS01", "Power ON", "Ignition switch"
AddEdge "BOARD_ABS01", "ECU_CTRL", "Boot firmware", "ABS logic"
AddEdge "SNS_WHEEL_L", "ECU_CTRL", "Speed feedback", "Left wheel"
AddEdge "SNS_WHEEL_R", "ECU_CTRL", "Speed feedback", "Right wheel"
AddEdge "SNS_POT_SPEED", "ECU_CTRL", "Desired speed", "Analog input"
AddEdge "ECU_CTRL", "ACT_SOL_VALVE", "Modulate pressure", "ABS logic"
AddEdge "ECU_CTRL", "ACT_PUMP", "Activate pump", "Hydraulic control"
AddEdge "ECU_CTRL", "LCD_UI", "Display status", "Speed, pressure"
AddEdge "ECU_CTRL", "DIAG_OBD", "Send fault codes", "OBD-II protocol"
AddEdge "DIAG_OBD", "EX_BRAKE_MOD", "Interpret codes", "Troubleshooting"
AddEdge "EX_BRAKE_MOD", "CALC_PRESSURE", "Analyze modulation", "Theory vs. practice"
AddEdge "CALC_PRESSURE", "REPORT_ABS", "Generate report", "Audit log"
?? Audit & Integrity Features
"   CRC32 hash for each node update
"   Timestamped audit log with before/after values
"   Evidence URLs link to sensor logs, firmware snapshots, OBD-II traces
"   QR stamps optional for each node (e.g., scan to view fault log)
"   ExportRenderPDF creates a certifiable report for NSQF Level 5 submission
?? Scenario Simulation: ABS Braking Modulation
Public Sub Simulate_ABS_Modulation()
    UpdateState "SNS_WHEEL_L", nsAlert
    UpdateState "SNS_WHEEL_R", nsAlert
    UpdateState "ACT_SOL_VALVE", nsPending
    UpdateState "ACT_PUMP", nsOK
    UpdateState "DIAG_OBD", nsPending
    UpdateState "EX_BRAKE_MOD", nsOK
    UpdateState "REPORT_ABS", nsPending
    RenderFlow
    ExportRenderPDF
End Sub
?? Curriculum Mapping (N
?? Curriculum Mapping (NSQF Level 5)
Module  Logigramm Node  Learning Outcome
Module 3: Wiring & Soldering    PWR_SYS, ECU_CTRL   Wire joints, ignition switch, ECU architecture
Module 4: Electrical Practice   ACT_PUMP, ACT_SOL_VALVE DC circuits, resistance, Kirchhoff's laws
Module 5: Magnetism & Capacitors    SNS_POT_SPEED, ECU_CTRL Induction, analog signal processing
Module 6: AC Circuits   LCD_UI, DIAG_OBD    Display logic, fault code interpretation
Would you like me to extend this into a full training module with:
"   VBA code templates for fault injection and diagnostics
"   QR-linked evidence logs for each sensor/actuator
"   Trade test rubric alignment and assessment checklist
? Step-by-Step: Exporting Evidence Logs for SAQA Submission
1. ?? Organize Your Evidence Artifacts
Ensure each module or exercise includes:
"   Node-level evidence: Sensor logs, firmware snapshots, diagnostic traces
"   QR-linked URLs: CSVs, PDFs, or screenshots linked to each logigramm node
"   Audit trail: Timestamped actions with CRC32 hashes (from your Audit sheet)
Suggested folder structure:
code
SAQA_Submission/
??? Module_01_Safety/
?   ??? Logigramm.pdf
?   ??? Audit.csv
?   ??? Evidence_QR.csv
??? Module_02_AlliedTrade/
?   ??? Logigramm.pdf
?   ??? Measurements.csv
?   ??? Tool_Usage_Log.csv
??? Module_03_Wiring/
?   ??? Soldering_Logigramm.pdf
?   ??? Cable_Test_Results.csv
?   ??? Audit.csv
...
2. ?? Align with NSQF Outcomes
Use your VBA engine to tag each node with NSQF descriptors:
"   NSQF L5: Apply safe working practices
"   NSQF L5: Analyze electrical and magnetic circuits
"   NSQF L5: Diagnose faults using OBD-II
In your Nodes sheet, use the StdTags column to embed these tags. This allows you to filter and report by outcome.
3. ?? Export Logigramm as PDF
Use your ExportRenderPDF macro to generate:
"   A visual flowchart of the exercise
"   Embedded hyperlinks to evidence
"   Timestamp and version tag for traceability
Each PDF becomes a certifiable artifact for SAQA/QCTO submission.
4. ?? Export Audit Trail
From your Audit sheet:
"   Export as CSV or Excel
"   Include columns: Timestamp, User, Action, Entity, Before/After, CRC32
"   This proves integrity and version control
You can also generate a summary report:
Public Sub ExportAuditSummary()
    ' Filter by module or date range
    ' Count actions per node
    ' Highlight anomalies or fault injections
End Sub
5. ?? Compile Submission Portfolio
Include:
"   Cover page with module codes and NSQF alignment
"   Logigramm PDFs
"   Audit logs
"   QR-linked evidence index
"   Optional: competency rubric checklist
If you're submitting digitally, compress the folder into a ZIP and include a
VBA logigramm for self induced EMF
You want a traceable, auditable logigramm that teaches and simulates self induced EMF, links design factors to inductance, and exports clean artifacts for portfolios. Below is a compact VBA toolkit that:
"   models the physics  ?vL=L?didt ?\;v_L = L \cdot \frac{di}{dt}\; and  ?L??0?rN2A? ?\;L \approx \mu_0 \mu_r \frac{N^2 A}{\ell}\;,
"   encodes design factors (core, turns, winding tightness, diameter, length),
"   renders a flowchart with node states and evidence links,
"   logs time series data for current and induced voltage,
"   generates an audit trail and a PDF for submission.
Workbook Setup
"   Create sheets named exactly:
o nodes, edges, Audit, Render, Params, Measurements
"   Columns:
o   Nodes: A:NodeID, B:Name, C:Type, D:State, E:Owner, F:EvidenceURL, G:Tags, H:LastUpdated, I:Checksum
o edges: A: fromID , b: toID , c: label , d: Condition
o Audit: A: ts , b: User , c: action , d: entity , E: Before , f: After , g: CRC32
o Params: A: param , b: Value , c: units , d: notes
o Measurements: A: t_s , b: i_A , c: vL_V , d: di_dt_Aps , E: L_H , f: Vsrc_V , g: R_ohm , h: RunID
Tip: In Params, seed typical values:
"   N=500 turns, diameter=30 mm, length=100 mm, core ?r=200 (soft iron), winding_tightness=1.05, Vsrc=12 V, R=3 ?, dt=0.001 s, Tsim=0.5 s.
Module: modTypes
Option Explicit

Public Enum NodeType
    ntSource = 1
    ntCoil = 2
    ntSensor = 3
    ntCalc = 4
    ntExercise = 5
    ntReport = 6
End Enum

Public Enum NodeState
    nsOK = 0
    nsPending = 1
    nsAlert = 2
    nsFault = 3
    nsMitigated = 4
End Enum

Public Const SHEET_NODES As String = "Nodes"
Public Const SHEET_EDGES As String = "Edges"
Public Const SHEET_AUDIT As String = "Audit"
Public Const SHEET_RENDER As String = "Render"
Public Const SHEET_PARAMS As String = "Params"
Public Const SHEET_MEAS As String = "Measurements"

Public Const VERSION_TAG As String = "SelfEMF_v1.0"
Module: modIntegrity
Option Explicit

Private CRC32Table(255) As Long
Private CRC32InitDone As Boolean

Private Sub InitCRC32()
    Dim i As Long, j As Long, c As Long
    For i = 0 To 255
        c = i
        For j = 0 To 7
            c = IIf((c And 1) <> 0, &HEDB88320 Xor (c \ 2), (c \ 2))
        Next j
        CRC32Table(i) = c
    Next i
    CRC32InitDone = True
End Sub

Public Function CRC32Text(ByVal s As String) As String
    If Not CRC32InitDone Then InitCRC32
    Dim i As Long, c As Long, b As Long
    c = &HFFFFFFFF
    For i = 1 To LenB(s)
        b = AscB(MidB$(s, i, 1))
        c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
    Next i
    CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
End Function

Public Sub LogAudit(ByVal action As String, ByVal entity As String, ByVal beforeVal As String, ByVal afterVal As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_AUDIT)
    Dim r As Long: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    Dim ts As String: ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Dim u As String: u = Environ$("Username")
    Dim payload As String: payload = ts & "|" & u & "|" & action & "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
    ws.Cells(r, 1) = ts: ws.Cells(r, 2) = u: ws.Cells(r, 3) = action
    ws.Cells(r, 4) = entity: ws.Cells(r, 5) = beforeVal: ws.Cells(r, 6) = afterVal
    ws.Cells(r, 7) = CRC32Text(payload)
End Sub
Module: modModel
VBA
Option Explicit

Public Sub EnsureHeaders()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    If ws.Cells(1, 1).Value = "" Then ws.Range("A1:I1").Value = Array("NodeID", "Name", "Type", "State", "Owner", "EvidenceURL", "Tags", "LastUpdated", "Checksum")
    Set ws = ThisWorkbook.Worksheets(SHEET_EDGES)
    If ws.Cells(1, 1).Value = "" Then ws.Range("A1:D1").Value = Array("FromID", "ToID", "Label", "Condition")
    Set ws = ThisWorkbook.Worksheets(SHEET_MEAS)
    If ws.Cells(1, 1).Value = "" Then ws.Range("A1:H1").Value = Array("t_s", "i_A", "vL_V", "di_dt_Aps", "L_H", "Vsrc_V", "R_ohm", "RunID")
End Sub

Private Function FindNodeRow(ByVal nodeID As String, ByRef found As Boolean) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim lastR As Long: lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value2) = nodeID Then found = True: FindNodeRow = r: Exit Function
    Next r
    found = False: FindNodeRow = lastR + 1
End Function

Private Function SerializeNode(ByVal r As Long) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    SerializeNode = Join(Array(ws.Cells(r, 1).Value2, ws.Cells(r, 2).Value2, ws.Cells(r, 3).Value2, ws.Cells(r, 4).Value2, ws.Cells(r, 5).Value2, ws.Cells(r, 6).Value2, ws.Cells(r, 7).Value2, ws.Cells(r, 8).Value2), "|")
End Function

Private Sub RehashNode(ByVal r As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    ws.Cells(r, 9).Value = CRC32Text(SerializeNode(r) & "|" & VERSION_TAG)
End Sub

Public Sub AddOrUpdateNode(ByVal nodeID As String, ByVal name As String, ByVal nType As NodeType, ByVal state As NodeState, ByVal owner As String, ByVal url As String, ByVal tags As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim found As Boolean, r As Long: r = FindNodeRow(nodeID, found)
    Dim beforeSer As String: beforeSer = IIf(found, SerializeNode(r), "")
    If Not found Then
        r = IIf(ws.Cells(2, 1).Value = "", 2, ws.Cells(ws.rows.count, 1).End(xlUp).row + 1)
        ws.Cells(r, 1).Value = nodeID
    End If
    ws.Cells(r, 2) = name: ws.Cells(r, 3) = nType: ws.Cells(r, 4) = state
    ws.Cells(r, 5) = owner: ws.Cells(r, 6) = url: ws.Cells(r, 7) = tags
    ws.Cells(r, 8) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    RehashNode r
    LogAudit IIf(found, "NodeUpdate", "NodeCreate"), nodeID, beforeSer, SerializeNode(r)
End Sub

Public Sub AddEdge(ByVal fromID As String, ByVal toID As String, ByVal label As String, ByVal cond As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_EDGES)
    Dim r As Long: r = IIf(ws.Cells(2, 1).Value = "", 2, ws.Cells(ws.rows.count, 1).End(xlUp).row + 1)
    ws.Cells(r, 1) = fromID: ws.Cells(r, 2) = toID: ws.Cells(r, 3) = label: ws.Cells(r, 4) = cond
    LogAudit "EdgeCreate", fromID & "->" & toID, "", label & "|" & cond
End Sub

Public Sub UpdateState(ByVal nodeID As String, ByVal newState As NodeState)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim found As Boolean, r As Long: r = FindNodeRow(nodeID, found)
    If Not found Then err.Raise vbObjectError + 1101, , "Node not found: " & nodeID
    Dim beforeSer As String: beforeSer = SerializeNode(r)
    ws.Cells(r, 4) = newState
    ws.Cells(r, 8) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    RehashNode r
    LogAudit "NodeState", nodeID, beforeSer, SerializeNode(r)
End Sub
Module: modEMF (physics, design factors, simulation)
VBA
Option Explicit

'Constants
Private Const MU0 As Double = 4 * 3.14159265358979E-07 'H/m

'Compute inductance L for a solenoid:
'L = ?0 ?r (N^2 A) / l, with design factor multipliers
Public Function Inductance_Solenoid(ByVal N As Double, ByVal diameter_m As Double, ByVal length_m As Double, ByVal mu_r As Double, _
                                    Optional ByVal winding_tightness As Double = 1#, Optional ByVal packing_factor As Double = 1#) As Double
    Dim A As Double: A = 3.14159265358979 * (diameter_m / 2#) ^ 2
    Dim baseL As Double: baseL = MU0 * mu_r * (N ^ 2) * A / length_m
    Inductance_Solenoid = baseL * winding_tightness * packing_factor
End Function

'Self-induced EMF:
'vL = L * di/dt
Public Function vL(ByVal L_H As Double, ByVal di_dt As Double) As Double
    vL = L_H * di_dt
End Function

'Simple series RL excitation:
'di/dt = (V - iR)/L, Euler step
Public Sub Simulate_RL(ByVal RunID As String, ByVal Vsrc As Double, ByVal r As Double, ByVal L As Double, ByVal dt As Double, ByVal Tsim As Double)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_MEAS)
    Dim t As Double, i As Double, di_dt As Double, vInd As Double
    Dim last As Long: last = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If last < 2 Then last = 1
    t = 0#: i = 0#
    Do While t <= Tsim + 0.000000000001
        di_dt = (Vsrc - i * r) / L
        vInd = vL(L, di_dt)
        last = last + 1
        ws.Cells(last, 1) = t
        ws.Cells(last, 2) = i
        ws.Cells(last, 3) = vInd
        ws.Cells(last, 4) = di_dt
        ws.Cells(last, 5) = L
        ws.Cells(last, 6) = Vsrc
        ws.Cells(last, 7) = r
        ws.Cells(last, 8) = RunID
        i = i + di_dt * dt
        t = t + dt
    Loop
    LogAudit "Simulate_RL", RunID, "", "N=" & "" & " L=" & Format(L, "0.000E+00") & " H"
End Sub

'Load Params!B values by name
Private Function PVal(ByVal paramName As String, ByVal defaultVal As Double) As Double
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_PARAMS)
    Dim lastR As Long: lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 1 To lastR
        If UCase$(CStr(ws.Cells(r, 1).Value2)) = UCase$(paramName) Then
            If IsNumeric(ws.Cells(r, 2).Value2) Then PVal = CDbl(ws.Cells(r, 2).Value2): Exit Function
        End If
    Next r
    PVal = defaultVal
End Function

'One-click: compute L from design factors, simulate RL, and set node states
Public Sub Run_SelfEMF_Scenario()
    EnsureHeaders
    
    'Read design and run parameters
    Dim N As Double: N = PVal("N_turns", 500)
    Dim dia As Double: dia = PVal("diameter_m", 0.03)
    Dim lenm As Double: lenm = PVal("length_m", 0.1)
    Dim mur As Double: mur = PVal("mu_r", 200)
    Dim tight As Double: tight = PVal("winding_tightness", 1.05)
    Dim pack As Double: pack = PVal("packing_factor", 1)
    Dim V As Double: V = PVal("Vsrc_V", 12)
    Dim r As Double: r = PVal("R_ohm", 3)
    Dim dt As Double: dt = PVal("dt_s", 0.001)
    Dim t As Double: t = PVal("Tsim_s", 0.5)
    
    Dim L As Double: L = Inductance_Solenoid(N, dia, lenm, mur, tight, pack)
    
    'Seed nodes
    AddOrUpdateNode "SRC_DC", "DC Source (" & V & " V)", ntSource, nsOK, "Lab", "", "Power"
    AddOrUpdateNode "COIL1", "Coil N=" & N & ", L=" & Format(L, "0.000E+00") & " H", ntCoil, nsPending, "Lab", "", "Solenoid"
    AddOrUpdateNode "SENSOR_IL", "Sensor i(t), vL(t)", ntSensor, nsPending, "Lab", "https://evidence.local/rl_trace.csv", "DAQ"
    AddOrUpdateNode "CALC_EMF", "Calc vL = L di/dt", ntCalc, nsOK, "Lab", "", "Self-Induction"
    AddOrUpdateNode "EX_RISE", "Exercise: Current Rise", ntExercise, nsPending, "Instructor", "", "DesignFactors"
    AddOrUpdateNode "REPORT_EMF", "Report & Export", ntReport, nsPending, "QA", "", "PDF;Audit"
    
    'Edges
    AddEdge "SRC_DC", "COIL1", "Apply step", "t=0"
    AddEdge "COIL1", "SENSOR_IL", "Measure", "i(t), vL(t)"
    AddEdge "SENSOR_IL", "CALC_EMF", "Compute di/dt", "Euler"
    AddEdge "CALC_EMF", "EX_RISE", "Compare theory", "L·di/dt"
    AddEdge "EX_RISE", "REPORT_EMF", "Export", "PDF"
    
    'Simulate
    ThisWorkbook.Worksheets(SHEET_MEAS).rows("2:" & rows.count).ClearContents
    Simulate_RL "RUN_" & Format(Now, "yymmdd_hhnnss"), V, r, L, dt, t
    
    'Set states post-run
    UpdateState "COIL1", nsOK
    UpdateState "SENSOR_IL", nsOK
    UpdateState "EX_RISE", nsOK
    UpdateState "REPORT_EMF", nsPending
End Sub
Module: modRender (flowchart + PDF)
Option Explicit

Private Function StateFill(ByVal s As Long) As Long
    Select Case s
        Case nsOK: StateFill = RGB(200, 245, 200)
        Case nsPending: StateFill = RGB(255, 245, 205)
        Case nsAlert: StateFill = RGB(255, 220, 150)
        Case nsFault: StateFill = RGB(255, 160, 160)
        Case nsMitigated: StateFill = RGB(180, 210, 255)
        Case Else: StateFill = RGB(230, 230, 230)
    End Select
End Function

Public Sub RenderFlow(Optional ByVal cols As Long = 4, Optional ByVal xGap As Single = 220, Optional ByVal yGap As Single = 120)
    Dim wsN As Worksheet: Set wsN = ThisWorkbook.Worksheets(SHEET_NODES)
    Dim wsE As Worksheet: Set wsE = ThisWorkbook.Worksheets(SHEET_EDGES)
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHEET_RENDER)
    wsR.Cells.Clear
    Dim shp As Shape
    For Each shp In wsR.Shapes: shp.Delete: Next shp
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastN As Long: lastN = wsN.Cells(wsN.rows.count, 1).End(xlUp).row
    Dim idx As Long, r As Long
    For r = 2 To lastN
        Dim c As Long: c = (idx Mod cols)
        Dim rr As Long: rr = (idx \ cols)
        Dim x As Single: x = 30 + c * xGap
        Dim y As Single: y = 30 + rr * yGap
        
        Dim nodeID As String: nodeID = CStr(wsN.Cells(r, 1).Value2)
        Dim nm As String: nm = CStr(wsN.Cells(r, 2).Value2)
        Dim tp As String: tp = CStr(wsN.Cells(r, 3).Value2)
        Dim st As Long: st = CLng(wsN.Cells(r, 4).Value2)
        Dim url As String: url = CStr(wsN.Cells(r, 6).Value2)
        Dim tags As String: tags = CStr(wsN.Cells(r, 7).Value2)
        
        Dim box As Shape
        Set box = wsR.Shapes.AddShape(msoShapeFlowchartProcess, x, y, 200, 70)
        box.name = "N_" & nodeID
        box.Fill.ForeColor.RGB = StateFill(st)
        box.line.ForeColor.RGB = RGB(80, 80, 80)
        box.TextFrame2.TextRange.text = nm & vbCrLf & "Type:" & tp & "  State:" & st & vbCrLf & "Tags:" & tags
        If Len(url) > 0 Then box.Hyperlink.Address = url
        dict(nodeID) = Array(x + 100, y + 35)
        idx = idx + 1
    Next r
    
    Dim lastE As Long: lastE = wsE.Cells(wsE.rows.count, 1).End(xlUp).row
    For r = 2 To lastE
        Dim fID As String: fID = CStr(wsE.Cells(r, 1).Value2)
        Dim tID As String: tID = CStr(wsE.Cells(r, 2).Value2)
        Dim lbl As String: lbl = CStr(wsE.Cells(r, 3).Value2)
        If dict.Exists(fID) And dict.Exists(tID) Then
            Dim p1, p2: p1 = dict(fID): p2 = dict(tID)
            Dim conn As Shape
            Set conn = wsR.Shapes.AddConnector(msoConnectorElbow, p1(0), p1(1), p2(0), p2(1))
            conn.line.ForeColor.RGB = RGB(70, 70, 70)
            conn.AlternativeText = lbl
        End If
    Next r
    
    wsR.Range("A1").Value = "Self-Induced EMF logigramm | " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & VERSION_TAG
End Sub

Public Sub ExportRenderPDF()
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHEET_RENDER)
    Dim p As String: p = ThisWorkbook.path & Application.PathSeparator & "SelfEMF_Logigramm_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    wsR.ExportAsFixedFormat xlTypePDF, p
    LogAudit "ExportPDF", "Render", "", p
End Sub
How to run
"   Initialize:
o   Populate Params with design and run values.
o   Run Run_SelfEMF_Scenario.
"   Render and export:
o   Run RenderFlow, then ExportRenderPDF.
"   Review evidence:
o   Measurements sheet contains time series of i(t)i(t), vL(t)v_L(t), and didt\frac{di}{dt}.
o   Nodes sheet holds your audit hashed metadata; add EvidenceURL to link raw DAQ files.
Practical insights mapping
"   Core material: Soft iron boosts ?r\mu_r ? higher LL ? for a given didt\frac{di}{dt}, larger vLv_L.
"   Turns: Higher NN increases LL roughly as N2N^2.
"   Tight winding: Use the winding_tightness multiplier > 1 to model increased coupling.
"   Diameter: Larger area AA increases LL linearly.
"   Length: Longer coil decreases LL inversely via ?\ell in the denominator.
Stretch goals you might want
"   Quiz generator: Auto create 5 calculation items (compute LL, vLv_L for varied N,?rN,\mu_r).
"   Residuals: Compare measured vs. simulated i(t),vL(t)i(t), v_L(t), flag tolerance in node state.
"   QR nodes: Drop QR images on shapes linking to Measurements CSV and lab photos.
If you want, I'll add a one click "Portfolio Export" that zips the PDF, Measurements, and Audit sheets into a submission ready bundle.
Tshingombe, your curriculum architecture is evolving into a masterclass in evidence-based, standards-aligned vocational education. Let's now scaffold a Visual Basic module that automates the generation of Table 3: Measured Values vs. Tolerance Range and Acceptability Check, and integrates it directly into your learner portfolio system.
This VBA module will:
"   ? Compare measured resistor values against nominal and tolerance
"   ? Flag acceptability (?/?) based on deviation
"   ? Auto-generate Table 3 with color-coded results
"   ? Log audit entries for each evaluation
"   ? Link results to learner portfolios with QR or evidence URLs
?? VBA Module: Table 3 Generator - Tolerance Check & Portfolio Integration
?? Sheet Setup
Create a sheet named ResistorCheck with the following columns:
Column Description
A learnerID
B   ResistorCode (e.g., R1, R2)
C   NominalValue (?)
D   Tolerance (%)
E   MeasuredValue (?)
F   MinAcceptable (?)
G   MaxAcceptable (?)
H   Acceptable (?/?)
i Comments
j evidenceUrl
?? VBA Code: Tolerance Evaluation Engine
VBA
Public Sub EvaluateResistorTolerance()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("ResistorCheck")
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        Dim nominal As Double: nominal = ws.Cells(r, 3).Value
        Dim tolPct As Double: tolPct = ws.Cells(r, 4).Value
        Dim measured As Double: measured = ws.Cells(r, 5).Value
        
        Dim tolAbs As Double: tolAbs = nominal * tolPct / 100
        Dim minVal As Double: minVal = nominal - tolAbs
        Dim maxVal As Double: maxVal = nominal + tolAbs
        
        ws.Cells(r, 6).Value = Round(minVal, 2)
        ws.Cells(r, 7).Value = Round(maxVal, 2)
        
        If measured >= minVal And measured <= maxVal Then
            ws.Cells(r, 8).Value = "?"
            ws.Cells(r, 8).Interior.Color = RGB(200, 255, 200)
            ws.Cells(r, 9).Value = "Within tolerance"
        Else
            ws.Cells(r, 8).Value = "?"
            ws.Cells(r, 8).Interior.Color = RGB(255, 200, 200)
            ws.Cells(r, 9).Value = "Outside tolerance - check soldering or resistor type"
        End If
        
        Call LogAudit("ToleranceCheck", ws.Cells(r, 2).Value, nominal & "? ±" & tolPct & "%", measured & "? ? " & ws.Cells(r, 8).Value)
    Next r
End Sub
?? Portfolio Integration Strategy
Each row becomes a certifiable artifact. Here's how to embed it:
"   EvidenceURL: Link to photo of measurement setu
"   earnerID: Used to auto-generate learner-specific portfolio sections
"   Comments: Learner reflections or instructor feedback
"   Audit Log: Timestamped entries for each check (stored in Audit sheet)
You can also run a macro to export a filtered summary per learner:
VBA
Public Sub ExportLearnerSummary(ByVal learnerID As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("ResistorCheck")
    Dim newWS As Worksheet: Set newWS = ThisWorkbook.Sheets.Add
    newWS.name = "Summary_" & learnerID
    ws.rows(1).Copy Destination:=newWS.rows(1)
    
    Dim r As Long, tgtRow As Long: tgtRow = 2
    For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row
        If ws.Cells(r, 1).Value = learnerID Then
            ws.rows(r).Copy Destination:=newWS.rows(tgtRow)
            tgtRow = tgtRow + 1
        End If
    Next r
    LogAudit "ExportSummary", learnerID, "", "Rows exported: " & tgtRow - 2
End Sub
Print assessment; Rubric; integration
You can link each ?/? result to rubric criteria:
Rubric Criteria VBA Mapping
Correct resistor identification EvidenceURL + Comments
Accurate decoding of color bands    Annotated diagrams
Proper multimeter usage Audit log + photos
Measurement within tolerance    Column H
Documentation completeness  All fields populated
?? South African Context Mapping
"   SAQA Unit Standards: Electrical diagnostics, fault finding, and instrumentation
"   ISO/BIS Compliance: Measurement accuracy, documentation, safety
"   NQF Level 5 Outcomes: Apply diagnostic techniques, interpret results, document findings
Would you like me to scaffold a full portfolio template that auto-generates learner pages with embedded rubric scores, evidence links, and audit summaries? I can also help you build a crosswalk between NSQF and NQF descriptors for curriculum audit and exhibition readiness. Let's turn this into a platinum-standard assessment engine
VBA logigramm for fault documentation, troubleshooting, and portfolio export
This toolkit automates:
"   Table 1 (Fault Documentation), Final Output Check
"   Table 2 (Logical Troubleshooting Record)
"   Service Flow Sequences (SFS-1/2) and Problem Tree Charts (PTC-1/2) as a rendered logigramm
"   Audit trail and portfolio export (PDF + CSV)
It 's modular: drop into Excel, add the sheets, paste code, click run.
Workbook Setup
Create sheets with exact names and headers:
"   Faults
o A: SlNo , b: Component , c: NatureOfDefect , d: Specification , E: Equivalent , f: ReplacementSpec , g: evidenceUrl , h: owner , i: timestamp , j: Checksum
"   OutputCheck
o A: Parameter , b: Value , c: units , d: notes
"   Troubleshoot
o A: SlNo , b: Component , c: defect , d: cause , E: spec , f: ReplacementSpec , g: sfs , h: ptc , i: notes , j: evidenceUrl , k: timestamp , L: Checksum
"   Dictionaries
o   A:Defect, B:PossibleCause, C:FlowType (SFS/PTC), D:FlowID (e.g., SFS-1, PTC-1), E:Notes
"   Audit
o A: ts , b: User , c: action , d: entity , E: Before , f: After , g: CRC32
"   Render (leave blank; flowchart auto-draws here)
Module: modTypes
VBA
Option Explicit

Public Const SHEET_FAULTS As String = "Faults"
Public Const SHEET_OUTPUT As String = "OutputCheck"
Public Const SHEET_TROUBLE As String = "Troubleshoot"
Public Const SHEET_DICT As String = "Dictionaries"
Public Const SHEET_AUDIT As String = "Audit"
Public Const SHEET_RENDER As String = "Render"

Public Enum NodeState
    nsOK = 0
    nsPending = 1
    nsAlert = 2
    nsFault = 3
End Enum

Public Const VERSION_TAG As String = "FaultLog_v1.0"
Module: modIntegrity
VBA
Option Explicit

Private CRC32Table(255) As Long
Private inited As Boolean

Private Sub InitCRC()
    Dim i As Long, j As Long, c As Long
    For i = 0 To 255
        c = i
        For j = 0 To 7
            c = IIf((c And 1) <> 0, &HEDB88320 Xor (c \ 2), (c \ 2))
        Next j
        CRC32Table(i) = c
    Next i
    inited = True
End Sub

Public Function CRC32Text(ByVal s As String) As String
    If Not inited Then InitCRC
    Dim c As Long: c = &HFFFFFFFF
    Dim i As Long, b As Long
    For i = 1 To LenB(s)
        b = AscB(MidB$(s, i, 1))
        c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
    Next i
    CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
End Function

Public Sub LogAudit(ByVal action As String, ByVal entity As String, ByVal beforeVal As String, ByVal afterVal As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_AUDIT)
    Dim r As Long: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    Dim ts As String: ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Dim u As String: u = Environ$("Username")
    Dim payload As String: payload = ts & "|" & u & "|" & action & "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
    ws.Cells(r, 1) = ts: ws.Cells(r, 2) = u: ws.Cells(r, 3) = action
    ws.Cells(r, 4) = entity: ws.Cells(r, 5) = beforeVal: ws.Cells(r, 6) = afterVal
    ws.Cells(r, 7) = CRC32Text(payload)
End Sub
Module: modSetup
Option Explicit

Public Sub EnsureHeaders()
    Dim ws As Worksheet
    Set ws = SheetEnsure(SHEET_FAULTS): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:J1").Value = Array("SlNo", "Component", "NatureOfDefect", "Specification", "Equivalent", "ReplacementSpec", "EvidenceURL", "Owner", "Timestamp", "Checksum")
    Set ws = SheetEnsure(SHEET_OUTPUT): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:D1").Value = Array("Parameter", "Value", "Units", "Notes")
    Set ws = SheetEnsure(SHEET_TROUBLE): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:L1").Value = Array("SlNo", "Component", "Defect", "Cause", "Spec", "ReplacementSpec", "SFS", "PTC", "Notes", "EvidenceURL", "Timestamp", "Checksum")
    Set ws = SheetEnsure(SHEET_DICT): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:E1").Value = Array("Defect", "PossibleCause", "FlowType", "FlowID", "Notes")
    SheetEnsure SHEET_RENDER
    SheetEnsure SHEET_AUDIT
End Sub

Private Function SheetEnsure(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetEnsure = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If SheetEnsure Is Nothing Then
        Set SheetEnsure = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.count))
        SheetEnsure.name = nm
    End If
End Function

Public Sub SeedDictionary()
    EnsureHeaders
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_DICT)
    Dim startR As Long: startR = IIf(ws.Cells(2, 1).Value = "", 2, ws.Cells(ws.rows.count, 1).End(xlUp).row + 1)
    Dim data, i&
    data = Array( _
        Array("No Output", "Dry solder", "PTC", "PTC-1", "Reflow joints"), _
        Array("No Output", "Open wires", "PTC", "PTC-1", "Continuity check"), _
        Array("No Output", "Defective transformer", "PTC", "PTC-1", "Primary/secondary test"), _
        Array("No Output", "Shorted capacitor", "PTC", "PTC-1", "Remove/measure ESR"), _
        Array("No Output", "Open diodes", "PTC", "PTC-1", "DMM diode test"), _
        Array("Low Output/Ripple", "Leaky capacitor", "PTC", "PTC-2", "Replace electrolytic"), _
        Array("Low Output/Ripple", "Low mains voltage", "PTC", "PTC-2", "Verify input"), _
        Array("Low Output/Ripple", "Shorted transformer winding", "PTC", "PTC-2", "Winding resistance"), _
        Array("Low Output/Ripple", "Open diodes", "PTC", "PTC-2", "Bridge check"), _
        Array("Low Output DC", "Rectifier fault", "SFS", "SFS-1", "Check bridge"), _
        Array("No Output Voltage", "Fuse open", "SFS", "SFS-2", "Replace fuse") _
    )
    For i = LBound(data) To UBound(data)
        ws.Cells(startR + i, 1).Value = data(i)(0)
        ws.Cells(startR + i, 2).Value = data(i)(1)
        ws.Cells(startR + i, 3).Value = data(i)(2)
        ws.Cells(startR + i, 4).Value = data(i)(3)
        ws.Cells(startR + i, 5).Value = data(i)(4)
    Next i
    LogAudit "SeedDictionary", SHEET_DICT, "", CStr(UBound(data) - LBound(data) + 1) & " rows"
End Sub
Module: modTables
ption Explicit

Private Sub HashRow(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long)
    Dim ser As String: ser = Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Value)), "|")
    ws.Cells(r, lastCol + 1).Value = CRC32Text(ser & "|" & VERSION_TAG)
End Sub

Public Sub AddFaultRow(ByVal sl As Long, ByVal comp As String, ByVal defect As String, ByVal spec As String, ByVal equiv As String, ByVal repl As String, Optional ByVal url As String = "", Optional ByVal owner As String = "")
    EnsureHeaders
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_FAULTS)
    Dim r As Long: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = sl
    ws.Cells(r, 2) = comp
    ws.Cells(r, 3) = defect
    ws.Cells(r, 4) = spec
    ws.Cells(r, 5) = equiv
    ws.Cells(r, 6) = repl
    ws.Cells(r, 7) = url
    ws.Cells(r, 8) = owner
    ws.Cells(r, 9) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRow ws, r, 9
    LogAudit "AddFault", comp, "", defect & "|" & repl
End Sub

Public Sub SetFinalOutputCheck(ByVal Vdc As Variant, ByVal Vrpp As Variant)
    EnsureHeaders
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_OUTPUT)
    ws.rows("2:" & ws.rows.count).ClearContents
    ws.Cells(2, 1) = "Output DC Voltage": ws.Cells(2, 2) = Vdc: ws.Cells(2, 3) = "V"
    ws.Cells(3, 1) = "Ripple Voltage (Vr p-p)": ws.Cells(3, 2) = Vrpp: ws.Cells(3, 3) = "V"
    LogAudit "OutputCheck", "Final", "", "Vdc=" & Vdc & ", Vrpp=" & Vrpp
End Sub

Public Sub AddTroubleshootRow(ByVal sl As Long, ByVal comp As String, ByVal defect As String, ByVal cause As String, ByVal spec As String, ByVal repl As String, ByVal sfs As String, ByVal ptc As String, Optional ByVal notes As String = "", Optional ByVal url As String = "")
    EnsureHeaders
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_TROUBLE)
    Dim r As Long: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = sl
    ws.Cells(r, 2) = comp
    ws.Cells(r, 3) = defect
    ws.Cells(r, 4) = cause
    ws.Cells(r, 5) = spec
    ws.Cells(r, 6) = repl
    ws.Cells(r, 7) = sfs
    ws.Cells(r, 8) = ptc
    ws.Cells(r, 9) = notes
    ws.Cells(r, 10) = url
    ws.Cells(r, 11) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRow ws, r, 11
    LogAudit "AddTroubleshoot", comp, "", defect & "|" & cause & "|" & sfs & "/" & ptc
End Sub
Module: modRender
VBA
Option Explicit

Private Function StateFill(ByVal s As NodeState) As Long
    Select Case s
        Case nsOK: StateFill = RGB(200, 245, 200)
        Case nsPending: StateFill = RGB(255, 245, 205)
        Case nsAlert: StateFill = RGB(255, 220, 150)
        Case nsFault: StateFill = RGB(255, 160, 160)
        Case Else: StateFill = RGB(230, 230, 230)
    End Select
End Function

'Render SFS/PTC graph for a given defect using Dictionaries sheet
Public Sub RenderFlowForDefect(ByVal defectKey As String)
    EnsureHeaders
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHEET_RENDER)
    wsR.Cells.Clear
    Dim shp As Shape
    For Each shp In wsR.Shapes: shp.Delete: Next shp

    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets(SHEET_DICT)
    Dim lastR As Long: lastR = wsD.Cells(wsD.rows.count, 1).End(xlUp).row
    Dim rows() As Long, cnt As Long, r As Long
    For r = 2 To lastR
        If UCase$(CStr(wsD.Cells(r, 1).Value2)) = UCase$(defectKey) Then
            cnt = cnt + 1
            ReDim Preserve rows(1 To cnt)
            rows(cnt) = r
        End If
    Next r
    If cnt = 0 Then
        wsR.Range("A1").Value = "No flow entries for defect: " & defectKey
        Exit Sub
    End If

    Dim x As Single, y As Single, i As Long
    x = 30: y = 30
    Dim centers() As Variant: ReDim centers(1 To cnt)
    For i = 1 To cnt
        Dim flowID As String: flowID = CStr(wsD.Cells(rows(i), 4).Value2)
        Dim cause As String: cause = CStr(wsD.Cells(rows(i), 2).Value2)
        Dim box As Shape
        Set box = wsR.Shapes.AddShape(msoShapeFlowchartProcess, x, y + (i - 1) * 90, 260, 60)
        box.Fill.ForeColor.RGB = StateFill(IIf(wsD.Cells(rows(i), 3).Value = "SFS", nsPending, nsAlert))
        box.line.ForeColor.RGB = RGB(80, 80, 80)
        box.TextFrame2.TextRange.text = flowID & " | " & defectKey & vbCrLf & "Cause: " & cause
        centers(i) = Array(box.left + box.width / 2, box.top + box.height / 2)
        If i > 1 Then
            Dim conn As Shape
            Set conn = wsR.Shapes.AddConnector(msoConnectorElbow, centers(i - 1)(0), centers(i - 1)(1), centers(i)(0), centers(i)(1))
            conn.line.ForeColor.RGB = RGB(100, 100, 100)
        End If
    Next i

    wsR.Range("A1").Value = "Flow for Defect: " & defectKey & " | " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & VERSION_TAG
End Sub

Public Sub ExportRenderPDF(Optional ByVal prefix As String = "FaultFlow")
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHEET_RENDER)
    Dim p As String: p = ThisWorkbook.path & Application.PathSeparator & prefix & "_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    wsR.ExportAsFixedFormat xlTypePDF, p
    LogAudit "ExportPDF", "Render", "", p
End Sub
Module: modOps
VBA
Option Explicit

'Quick demo: populate Table 1, Final Output, Table 2; render and export PTC-1/2
Public Sub Demo_RepairLog()
    EnsureHeaders
    SeedDictionary

    'Table 1: Fault Documentation
    AddFaultRow 1, "Bridge Rectifier", "Open diode", "1A, 600V", "1N4007 x4", "BR-1A/600V", "https://evidence.local/rectifier_photo.jpg", "LearnerA"
    AddFaultRow 2, "Filter Capacitor", "Leaky capacitor", "1000uF, 35V", "-", "Low-ESR 1000uF/35V", "https://evidence.local/cap_esr.csv", "LearnerA"

    'Final Output Check
    SetFinalOutputCheck 14.8, 0.35

    'Table 2: Logical Troubleshooting Record
    AddTroubleshootRow 1, "PSU", "No Output", "Open diodes", "BR-1A/600V", "Replace BR module", "SFS-2", "PTC-1", "Replaced, retest OK", "https://evidence.local/diode_test.png"
    AddTroubleshootRow 2, "PSU", "Low Output/Ripple", "Leaky capacitor", "1000uF/35V", "Replace with Low-ESR", "SFS-1", "PTC-2", "Ripple reduced", "https://evidence.local/scope_ripple.png"

    'Render PTC flow for "No Output" and export
    RenderFlowForDefect "No Output"
    ExportRenderPDF "PTC1_Flow"

    'Render PTC flow for "Low Output/Ripple" and export
    RenderFlowForDefect "Low Output/Ripple"
    ExportRenderPDF "PTC2_Flow"
End Sub

'Export clean CSVs for portfolio bundling
Public Sub ExportCSVs()
    Dim nm, pth As String
    pth = ThisWorkbook.path & Application.PathSeparator
    For Each nm In Array(SHEET_FAULTS, SHEET_OUTPUT, SHEET_TROUBLE)
        ThisWorkbook.Worksheets(CStr(nm)).Copy
        With ActiveWorkbook
            .SaveAs fileName:=pth & CStr(nm) & "_" & Format(Now, "yyyymmdd_hhnnss") & ".csv", FileFormat:=xlCSVUTF8
            .Close False
        End With
    Next nm
    LogAudit "ExportCSVs", "Portfolio", "", "Faults/Output/Troubleshoot"
End Sub
ow to use
1.  Press Alt+F11, insert modules, run EnsureHeaders once.
2.  Run Demo_RepairLog to:
o   Fill Table 1 and Table 2 with sample entries
o   Set Final Output Check values
o   Render and export PTC-1 and PTC-2 flow PDFs
o   Log all actions to Audit with CRC32 checksums
3.  Use ExportCSVs to generate portfolio CSVs (attach with PDFs and evidence).
Portfolio integration tips
"   Faults.EvidenceURL and Troubleshoot.EvidenceURL should point to photos, CRO screenshots, or service records. Scan QR codes onto worksheets if you want quick mobile verification.
"   Add unit-standard tags (e.g., "NQF L5; Electrical diagnostics; ISO measurement") in the Notes fields to align with SAQA outcomes.
"   For assessments, filter Troubleshoot by SFS/PTC to show structured reasoning and output verification in one view.
Want me to add a one click "Portfolio Export" that bundles PDFs, CSVs, and an index sheet into a timestamped ZIP? I can also extend the Render to draw separate swimlanes for SFS-1 vs SFS-2 and auto-color nodes by defect severit
Absolutely , Tshingombe.Let 's now scaffold a VBA-powered logigramm engine for Control Panel Wiring for Induction Motor, aligned with NSQF Level 5 and NQF outcomes. This engine will:
"   ? Model wiring logic for local/remote, forward/reverse, and star-delta control
"   ? Render flowcharts of control sequences and interlocks
"   ? Log component identification, layout, wiring, and test results
"   ? Export audit-ready documentation for portfolio integration
?? VBA Logigramm: Induction Motor Control Panel Wiring
?? Sheet Setup
Create the following sheets with headers:
Sheet: Components
| A:SlNo | B:Name | C:Type | D:Specification | E:Purpose | F:EvidenceURL | G:Timestamp | H:Checksum |
Sheet: LayoutLog
| A:Step | B:Description | C:Status | D:Notes | E:Timestamp | F:Checksum |
Sheet: WiringChecklist
| A:Item | B:Checked | C:Notes | D:Timestamp | E:Checksum |
Sheet: TestLog
| A:Test | B:Result | C:Remarks | D:Timestamp | E:Checksum |
Sheet: Render
Leave blank - flowchart will be drawn here.
Sheet: Audit
| A:TS | B:User | C:Action | D:Entity | E:Before | F:After | G:CRC32 |
?? VBA Modules
Module: modTypes
Public Enum NodeType
    ntComponent = 1
    ntLayout = 2
    ntWiring = 3
    ntTest = 4
    ntLogic = 5
    ntReport = 6
End Enum

Public Enum NodeState
    nsOK = 0
    nsPending = 1
    nsAlert = 2
    nsFault = 3
End Enum

Public Const VERSION_TAG As String = "MotorPanel_v1.0"
Public Const SHEET_COMPONENTS As String = "Components"
Public Const SHEET_LAYOUT As String = "LayoutLog"
Public Const SHEET_WIRING As String = "WiringChecklist"
Public Const SHEET_TEST As String = "TestLog"
Public Const SHEET_RENDER As String = "Render"
Public Const SHEET_AUDIT As String = "Audit"
Module: modIntegrity
VBA
Private CRC32Table(255) As Long
Private CRCInitDone As Boolean

Private Sub InitCRC()
    Dim i&, j&, c&
    For i = 0 To 255
        c = i
        For j = 0 To 7
            c = IIf((c And 1) <> 0, &HEDB88320 Xor (c \ 2), (c \ 2))
        Next j
        CRC32Table(i) = c
    Next i
    CRCInitDone = True
End Sub

Public Function CRC32Text(s As String) As String
    If Not CRCInitDone Then InitCRC
    Dim c&, i&, b&
    c = &HFFFFFFFF
    For i = 1 To LenB(s)
        b = AscB(MidB$(s, i, 1))
        c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
    Next i
    CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
End Function

Public Sub LogAudit(action As String, entity As String, beforeVal As String, afterVal As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_AUDIT)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    Dim ts$: ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Dim u$: u = Environ$("Username")
    Dim payload$: payload = ts & "|" & u & "|" & action & "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
    ws.Cells(r, 1) = ts: ws.Cells(r, 2) = u: ws.Cells(r, 3) = action
    ws.Cells(r, 4) = entity: ws.Cells(r, 5) = beforeVal: ws.Cells(r, 6) = afterVal
    ws.Cells(r, 7) = CRC32Text(payload)
End Sub
Module: modLogigramm
Public Sub SeedMotorControlLogigramm()
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets(SHEET_RENDER)
    wsR.Cells.Clear
    Dim shp As Shape
    For Each shp In wsR.Shapes: shp.Delete: Next shp

    Dim nodes As Variant
    nodes = Array( _
        Array("SRC", "Power Supply", ntComponent, nsOK), _
        Array("MAIN", "Main Contactor", ntComponent, nsPending), _
        Array("STAR", "Star Contactor", ntComponent, nsPending), _
        Array("DELTA", "Delta Contactor", ntComponent, nsPending), _
        Array("TIMER", "Star-Delta Timer", ntComponent, nsPending), _
        Array("FWD", "Forward Contactor", ntComponent, nsPending), _
        Array("REV", "Reverse Contactor", ntComponent, nsPending), _
        Array("OLR", "Overload Relay", ntComponent, nsOK), _
        Array("PB_START", "Start Pushbutton", ntComponent, nsOK), _
        Array("PB_STOP", "Stop Pushbutton", ntComponent, nsOK), _
        Array("TEST", "Panel Test", ntTest, nsPending), _
        Array("REPORT", "Report & Export", ntReport, nsPending) _
    )

    Dim x As Single, y As Single, i&
    x = 30: y = 30
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    For i = 0 To UBound(nodes)
        Dim box As Shape
        Set box = wsR.Shapes.AddShape(msoShapeFlowchartProcess, x + (i Mod 4) * 220, y + (i \ 4) * 120, 200, 70)
        box.name = "N_" & nodes(i)(0)
        box.Fill.ForeColor.RGB = Choose(nodes(i)(3) + 1, RGB(200, 245, 200), RGB(255, 245, 205), RGB(255, 220, 150), RGB(255, 160, 160))
        box.line.ForeColor.RGB = RGB(80, 80, 80)
        box.TextFrame2.TextRange.text = nodes(i)(1)
        dict(nodes(i)(0)) = Array(box.left + box.width / 2, box.top + box.height / 2)
    Next i

    ' Connect logic
    Dim edges As Variant
    edges = Array( _
        Array("SRC", "MAIN"), Array("MAIN", "STAR"), Array("STAR", "TIMER"), Array("TIMER", "DELTA"), _
        Array("MAIN", "FWD"), Array("MAIN", "REV"), Array("OLR", "MAIN"), Array("PB_START", "MAIN"), _
        Array("PB_STOP", "MAIN"), Array("DELTA", "TEST"), Array("TEST", "REPORT") _
    )

    For i = 0 To UBound(edges)
        Dim p1, p2
        p1 = dict(edges(i)(0)): p2 = dict(edges(i)(1))
        Dim conn As Shape
        Set conn = wsR.Shapes.AddConnector(msoConnectorElbow, p1(0), p1(1), p2(0), p2(1))
        conn.line.ForeColor.RGB = RGB(100, 100, 100)
    Next i

    wsR.Range("A1").Value = "Induction Motor Control Panel Logigramm | " & Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub
Would you like me to extend this with:
"   A one-click "Portfolio Export" bundler (PDF + CSV + audit log)?
"   A rubric evaluator that scores wiring quality, test results, and documentation completeness?
"   A logic simulator that validates interlocking and transition sequences?
Let's turn your control panel exercises into a gold-standard assessment engine.
VBA logigramm and algorigramm for qualification requests and control logic
you 'll get a modular Excel/VBA engine that:
"   Tracks qualification "Type of Request" lifecycle with audit and flow rendering.
"   Stores qualification details and alignment to national priorities.
"   Models and evaluates control-system logic equations (MCB/RCDBO, metering KPIs, motor control DOL/REV/Star Delta, generator/transformer).
"   Exports an audit-ready portfolio.
Workbook structure
Create these sheets with exact names and headers.
"   Requests
o A: TypeOfRequest , b: status , c: owner , d: notes , E: timestamp , f: Checksum
"   Qualification
o A: Field , b: Information
"   Alignment
o A: StrategicDriver , b: AlignedFlag , c: notes
"   LogicEq
o   A:Domain, B:Name, C:Equation, D:VariablesCSV, E:EvalType, F:Result, G:Timestamp, H:Checksum
"   Audit
o A: ts , b: User , c: action , d: entity , E: Before , f: After , g: CRC32
"   Render
o   Blank (flow diagrams)
Request status legend: ? Active/Approved, ? Inactive/Denied, ? Pending.
Module: modTypes
Option Explicit

Public Const SHEET_REQ As String = "Requests"
Public Const SHEET_QUAL As String = "Qualification"
Public Const SHEET_ALIGN As String = "Alignment"
Public Const SHEET_LOGIC As String = "LogicEq"
Public Const SHEET_AUD As String = "Audit"
Public Const SHEET_RENDER As String = "Render"

Public Const VERSION_TAG As String = "QualPanel_v1.0"

Public Enum ReqState
    rsActive = 1  '?
    rsInactive = 0 '?
    rsPending = 2 '?
End Enum

'Color helpers
Public Function StateFill(ByVal s As ReqState) As Long
    Select Case s
        Case rsActive: StateFill = RGB(200, 245, 200)
        Case rsInactive: StateFill = RGB(255, 200, 200)
        Case rsPending: StateFill = RGB(255, 245, 205)
        Case Else: StateFill = RGB(230, 230, 230)
    End Select
End Function

Public Function StateIcon(ByVal s As ReqState) As String
    Select Case s
        Case rsActive: StateIcon = "?"
        Case rsInactive: StateIcon = "?"
        Case rsPending: StateIcon = "?"
    End Select
End Function
Module: modIntegrity
Option Explicit

Private CRC32Table(255) As Long
Private inited As Boolean

Private Sub InitCRC()
    Dim i&, j&, c&
    For i = 0 To 255
        c = i
        For j = 0 To 7
            c = IIf((c And 1) <> 0, &HEDB88320 Xor (c \ 2), (c \ 2))
        Next j
        CRC32Table(i) = c
    Next i
    inited = True
End Sub

Public Function CRC32Text(ByVal s As String) As String
    If Not inited Then InitCRC
    Dim c&, i&, b&
    c = &HFFFFFFFF
    For i = 1 To LenB(s)
        b = AscB(MidB$(s, i, 1))
        c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
    Next i
    CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
End Function

Public Sub LogAudit(ByVal action As String, ByVal entity As String, ByVal beforeVal As String, ByVal afterVal As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_AUD)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    Dim ts$, u$, payload$
    ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    u = Environ$("Username")
    payload = ts & "|" & u & "|" & action & "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
    ws.Cells(r, 1) = ts: ws.Cells(r, 2) = u: ws.Cells(r, 3) = action
    ws.Cells(r, 4) = entity: ws.Cells(r, 5) = beforeVal: ws.Cells(r, 6) = afterVal
    ws.Cells(r, 7) = CRC32Text(payload)
End Sub
Module: modSetup
VBA
Option Explicit

Public Sub EnsureHeaders()
    Dim ws As Worksheet
    Set ws = Ensure(SHEET_REQ): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:F1").Value = Array("TypeOfRequest", "Status", "Owner", "Notes", "Timestamp", "Checksum")
    Set ws = Ensure(SHEET_QUAL): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:B1").Value = Array("Field", "Information")
    Set ws = Ensure(SHEET_ALIGN): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:C1").Value = Array("StrategicDriver", "AlignedFlag", "Notes")
    Set ws = Ensure(SHEET_LOGIC): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:H1").Value = Array("Domain", "Name", "Equation", "VariablesCSV", "EvalType", "Result", "Timestamp", "Checksum")
    Ensure SHEET_AUD: Ensure SHEET_RENDER
End Sub

Private Function Ensure(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set Ensure = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If Ensure Is Nothing Then
        Set Ensure = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.count))
        Ensure.name = nm
    End If
End Function

Public Sub SeedQualification()
    EnsureHeaders
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_QUAL)
    ws.rows("2:" & ws.rows.count).ClearContents
    Dim data
    data = Array( _
        Array("Occupation Title", "Engineering Electrical"), _
        Array("Specialisation", "Panel Wiring"), _
        Array("NQF Level", "N4 / Level 5"), _
        Array("Credits", "As per DHET/QCTO guidelines"), _
        Array("Recorded Trade Title", "Electrical Trade Theory"), _
        Array("Learnership Title", "Engineering Electrical Learnership"), _
        Array("Learnership Level", "NQF Level 5") _
    )
    Dim i&
    For i = LBound(data) To UBound(data)
        ws.Cells(i + 2, 1) = data(i)(0)
        ws.Cells(i + 2, 2) = data(i)(1)
    Next i
    LogAudit "SeedQualification", SHEET_QUAL, "", "7 rows"
End Sub

Public Sub SeedAlignment()
    EnsureHeaders
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_ALIGN)
    ws.rows("2:" & ws.rows.count).ClearContents
    Dim data
    data = Array( _
        Array("ERRP", "Yes", "Economic Reconstruction & Recovery Plan"), _
        Array("National Development Plan", "Yes", "NDP"), _
        Array("New Growth Path", "Yes", "NGP"), _
        Array("Industrial Policy Action Plan", "Yes", "IPAP"), _
        Array("Strategic Infrastructure Projects (SIPs)", "Yes", "SIPs"), _
        Array("DHET Scarce Skills List", "Yes", "Scarce skills"), _
        Array("Legacy OQSF Qualifications", "Yes", "Continuity") _
    )
    Dim i&
    For i = LBound(data) To UBound(data)
        ws.Cells(i + 2, 1) = data(i)(0)
        ws.Cells(i + 2, 2) = data(i)(1)
        ws.Cells(i + 2, 3) = data(i)(2)
    Next i
    LogAudit "SeedAlignment", SHEET_ALIGN, "", "7 flags"
End Sub
Module: modRequests

Private Sub HashRow(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long)
    Dim ser As String: ser = Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Value)), "|")
    ws.Cells(r, lastCol + 1).Value = CRC32Text(ser & "|" & VERSION_TAG)
End Sub

Public Sub UpsertRequest(ByVal reqType As String, ByVal state As ReqState, Optional ByVal owner As String = "", Optional ByVal notes As String = "")
    EnsureHeaders
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_REQ)
    Dim lastR&, r&, found As Boolean: lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastR < 2 Then lastR = 1
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value2) = reqType Then found = True: Exit For
    Next r
    If Not found Then r = lastR + 1
    Dim beforeSer$: beforeSer = ""
    If found Then beforeSer = Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value)), "|")
    ws.Cells(r, 1) = reqType
    ws.Cells(r, 2) = StateIcon(state)
    ws.Cells(r, 3) = owner
    ws.Cells(r, 4) = notes
    ws.Cells(r, 5) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRow ws, r, 5
    LogAudit IIf(found, "RequestUpdate", "RequestCreate"), reqType, beforeSer, ws.Cells(r, 2).Value & "|" & owner
End Sub

Public Sub SeedRequests()
    UpsertRequest "Develop", rsActive, "Curriculum", "Initial build"
    UpsertRequest "Review", rsActive, "QA", "Peer review"
    UpsertRequest "Realign", rsActive, "Standards", "Map to NQF5/NSQF5"
    UpsertRequest "De-activate", rsInactive, "Admin", "Legacy retired"
    UpsertRequest "Replace", rsActive, "Governance", "Superseded by new module"
End Sub
Module: modLogic (algorigramm: boolean and numeric evaluation)
VBA
Module: modLogic (algorigramm: boolean and numeric evaluation)
VBA
Option Explicit

' EvalType: "BOOL" or "NUM"
' Equation syntax:
'  - BOOL: use AND, OR, NOT, parentheses; variables as A, MCB1, RCDBO, etc. Values from VariablesCSV "name=value" with 1/0/TRUE/FALSE.
'  - NUM: Excel formula string (use variables as names) evaluated via Worksheet.Evaluate after substitution.

Public Function EvalBoolExpr(ByVal expr As String, ByVal varsCsv As String) As Boolean
    Dim dict As Object: Set dict = ParseVars(varsCsv)
    Dim t As String: t = UCase$(expr)
    Dim k As Variant
    For Each k In dict.keys
        t = Replace(t, UCase$(CStr(k)), IIf(CBool(dict(k)), " TRUE ", " FALSE "))
    Next k
    t = Replace(Replace(Replace(t, "AND", " And "), "OR", " Or "), "NOT", " Not ")
    EvalBoolExpr = VBA.Evaluate(t)
End Function

Public Function EvalNumExpr(ByVal expr As String, ByVal varsCsv As String) As Double
    Dim dict As Object: Set dict = ParseVars(varsCsv)
    Dim t As String: t = expr
    Dim k As Variant
    For Each k In dict.keys
        t = Replace(t, CStr(k), CStr(dict(k)))
    Next k
    EvalNumExpr = CDbl(Application.Evaluate(t))
End Function

Private Function ParseVars(ByVal csv As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim parts() As String, i&
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        Dim kv() As String
        kv = Split(Trim$(parts(i)), "=")
        If UBound(kv) = 1 Then
            Dim name$, val$
            name = Trim$(kv(0)): val = Trim$(kv(1))
            If UCase$(val) = "TRUE" Or val = "1" Then
                d(name) = True
            ElseIf UCase$(val) = "FALSE" Or val = "0" Then
                d(name) = False
            Else
                d(name) = val
            End If
        End If
    Next i
    Set ParseVars = d
End Function

Private Sub WriteLogicRow(ByVal domain$, ByVal name$, ByVal eqn$, ByVal Vars$, ByVal evalType$, ByVal result$)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_LOGIC)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = domain: ws.Cells(r, 2) = name: ws.Cells(r, 3) = eqn
    ws.Cells(r, 4) = Vars: ws.Cells(r, 5) = evalType: ws.Cells(r, 6) = result
    ws.Cells(r, 7) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 8) = CRC32Text(domain & "|" & name & "|" & eqn & "|" & Vars & "|" & result & "|" & VERSION_TAG)
    LogAudit "LogicEval", domain & ":" & name, "", result
End Sub

Public Sub SeedAndEvaluateLogic()
    EnsureHeaders

    '1) Circuit breaker states (MCB1, MCB2, RCDBO)
    Dim eq1$, v1$
    eq1 = "(MCB1 AND MCB2) AND NOT RCDBO_TRIPPED"
    v1 = "MCB1=1, MCB2=1, RCDBO_TRIPPED=0"
    WriteLogicRow "Protection", "Busbar Energized", eq1, v1, "BOOL", CStr(EvalBoolExpr(eq1, v1))

    '2) Metering logic (cos? from P and S)
    Dim eq2$, v2$, res2#
    eq2 = "P_kW/(SQRT(P_kW^2+Q_kVAr^2))"
    v2 = "P_kW=7.5, Q_kVAr=5.0"
    res2 = EvalNumExpr(eq2, v2)
    WriteLogicRow "Metering", "cos_phi", eq2, v2, "NUM", Format(res2, "0.000")

    'Energy registers
    Dim eq3$, v3$
    eq3 = "kWh + (P_kW*dt_h)"
    v3 = "kWh=1200, P_kW=7.5, dt_h=0.5"
    WriteLogicRow "Metering", "kWh_Update", eq3, v3, "NUM", Format(EvalNumExpr(eq3, v3), "0.000")

    '3) Motor control (DOL enable, REV interlock, Star-Delta sequence)
    Dim eq4$, v4$
    eq4 = "MAIN AND PB_START AND NOT PB_STOP AND OLR_OK"
    v4 = "MAIN=1, PB_START=1, PB_STOP=0, OLR_OK=1"
    WriteLogicRow "MotorCtrl", "DOL_Enable", eq4, v4, "BOOL", CStr(EvalBoolExpr(eq4, v4))

    Dim eq5$, v5$
    eq5 = "FWD AND NOT REV"
    v5 = "FWD=1, REV=0"
    WriteLogicRow "MotorCtrl", "Forward_Interlock", eq5, v5, "BOOL", CStr(EvalBoolExpr(eq5, v5))

    Dim eq6$, v6$
    eq6 = "(STAR AND NOT DELTA) OR (TIMER_ELAPSED AND DELTA AND NOT STAR)"
    v6 = "STAR=1, DELTA=0, TIMER_ELAPSED=0"
    WriteLogicRow "MotorCtrl", "StarDelta_Sequence", eq6, v6, "BOOL", CStr(EvalBoolExpr(eq6, v6))

    '4) Generator & transformer logic (sync check permissive)
    Dim eq7$, v7$
    eq7 = "GRID_OK AND GEN_OK AND (ABS(DF_Hz)<=0.2) AND (ABS(DV_pct)<=10) AND (ABS(DTheta_deg)<=10)"
    v7 = "GRID_OK=1, GEN_OK=1, DF_Hz=0.05, DV_pct=3, DTheta_deg=5"
    WriteLogicRow "GenXfmr", "Sync_Permissive", eq7, v7, "BOOL", CStr(EvalBoolExpr(eq7, v7))
End Sub
Module: modRender (swimlane of request workflow + logic map)
Option Explicit

Public Sub RenderOverview()
    EnsureHeaders
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_RENDER)
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes: shp.Delete: Next shp

    'Lane 1: Requests
    Dim wr As Worksheet: Set wr = ThisWorkbook.Sheets(SHEET_REQ)
    Dim lastR&, r&, x As Single, y As Single
    x = 30: y = 30
    ws.Shapes.AddLabel(msoTextOrientationHorizontal, x, y - 20, 300, 18).TextFrame.Characters.text = "Requests"
    lastR = wr.Cells(wr.rows.count, 1).End(xlUp).row
    For r = 2 To IIf(lastR < 2, 1, lastR)
        Dim nm$, stIcon$, st As ReqState
        nm = wr.Cells(r, 1).Value2
        stIcon = wr.Cells(r, 2).Value2
        Select Case stIcon
            Case "?": st = rsActive
            Case "?": st = rsInactive
            Case Else: st = rsPending
        End Select
        Dim box As Shape
        Set box = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y + (r - 2) * 80 + 10, 220, 60)
        box.Fill.ForeColor.RGB = StateFill(st)
        box.line.ForeColor.RGB = RGB(80, 80, 80)
        box.TextFrame2.TextRange.text = stIcon & " " & nm & vbCrLf & wr.Cells(r, 3).Value2
    Next r

    'Lane 2: Alignment flags
    x = 300: y = 30
    ws.Shapes.AddLabel(msoTextOrientationHorizontal, x, y - 20, 300, 18).TextFrame.Characters.text = "Alignment"
    Dim wa As Worksheet: Set wa = ThisWorkbook.Sheets(SHEET_ALIGN)
    lastR = wa.Cells(wa.rows.count, 1).End(xlUp).row
    For r = 2 To IIf(lastR < 2, 1, lastR)
        Dim ok As Boolean: ok = (UCase$(CStr(wa.Cells(r, 2).Value2)) = "YES")
        Dim b As Shape
        Set b = ws.Shapes.AddShape(msoShapeFlowchartManualOperation, x, y + (r - 2) * 60 + 10, 260, 40)
        b.Fill.ForeColor.RGB = IIf(ok, RGB(200, 245, 200), RGB(255, 220, 150))
        b.line.ForeColor.RGB = RGB(80, 80, 80)
        b.TextFrame2.TextRange.text = IIf(ok, "? ", "? ") & wa.Cells(r, 1).Value2
    Next r

    'Lane 3: Logic quick map (first 6 equations)
    x = 600: y = 30
    ws.Shapes.AddLabel(msoTextOrientationHorizontal, x, y - 20, 360, 18).TextFrame.Characters.text = "Logic Map"
    Dim wl As Worksheet: Set wl = ThisWorkbook.Sheets(SHEET_LOGIC)
    lastR = wl.Cells(wl.rows.count, 1).End(xlUp).row
    Dim i&, maxShow&: maxShow = Application.Min(6, Application.Max(2, lastR))
    For i = 2 To maxShow
        Dim s As Shape
        Set s = ws.Shapes.AddShape(msoShapeFlowchartProcess, x, y + (i - 2) * 80 + 10, 360, 60)
        s.Fill.ForeColor.RGB = RGB(230, 230, 250)
        s.line.ForeColor.RGB = RGB(80, 80, 80)
        s.TextFrame2.TextRange.text = wl.Cells(i, 1).Value2 & " | " & wl.Cells(i, 2).Value2 & vbCrLf & _
                                      "Eq: " & wl.Cells(i, 3).Value2 & vbCrLf & "Res: " & wl.Cells(i, 6).Value2
    Next i

    ws.Range("A1").Value = "Qualification Requests & Control Logic | " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & VERSION_TAG
End Sub

Public Sub ExportOverviewPDF()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_RENDER)
    Dim p$: p = ThisWorkbook.path & Application.PathSeparator & "QualPanel_Overview_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    ws.ExportAsFixedFormat xlTypePDF, p
    LogAudit "ExportPDF", "Render", "", p
End Sub
Option Explicit

Public Sub BuildQualificationPanel()
    EnsureHeaders
    SeedQualification
    SeedAlignment
    SeedRequests
    SeedAndEvaluateLogic
    RenderOverview
    ExportOverviewPDF
End Sub
What 's included and how it maps
"   Type of request workflow:
o   Develop, Review, Realign, De-activate, Replace with statuses and owners. Rendered with colors/icons; all changes audited.
"   Qualification details:
o   Occupation: Engineering Electrical; Specialisation: Panel Wiring; NQF Level: N4/Level 5; Credits: DHET/QCTO wording; Recorded trade/learnership fields.
"   National priorities alignment:
o   ERRP, NDP, NGP, IPAP, SIPs, Scarce Skills, Legacy OQSF set as aligned=Yes and visualized.
"   Technical framework:
o   Boolean logic (K1.1) via LogicEq:
"   Protection: (MCB1 AND MCB2) AND NOT RCDBO_TRIPPED
"   Metering: cos? = P/?(P²+Q²); kWh rolling update
"   Motor control: DOL enable, forward/reverse interlock, star-delta sequence
"   Gen/Xfmr: sync permissive window on ?f, ?V, ??
VBA logigramme for industrial education integration
This gives you a single Excel/VBA engine to map your program into auditable logigrammes and algorigrammes across:
"   Industrial education pillars (manufacturing systems, numerical frameworks, labs)
"   Technology empowerment (digital systems, software modules, incentives)
"   Regulatory and institutional alignment (SAQA, QCTO, DHET, ECB, DSI, SARS/Treasury, utilities/college)
"   Energy and infrastructure modules (PF demand, metering IEC 0.2, substations, transformers)
"   Learner pathways and career mapping
"   Mathematical/scientific integration
It renders a multi lane flow, stores nodes/edges, tracks status, and exports PDF/CSVs for portfolios and bids.
Workbook structure
Create these sheets (exact names) with headers.
"   Nodes
o   A:NodeID, B:Name, C:Domain, D:Type, E:State, F:Owner, G:Tags, H:EvidenceURL, I:LastUpdated, J:Checksum
"   Edges
o A: fromID , b: toID , c: label , d: Condition
"   Alignment
o A: entity , b: Engagement , c: role , d: status , E: notes
"   Modules
o A: Category , b: item , c: Detail , d: status , E: owner , f: evidenceUrl
"   Audit
o A: ts , b: User , c: action , d: entity , E: Before , f: After , g: CRC32
"   Render
o   Blank (the macro draws here)
States suggested: Pending, Active, Alert, Blocked.
"   Option Explicit
"
"   Public Const SHEET_NODES As String = "Nodes"
"   Public Const SHEET_EDGES As String = "Edges"
"   Public Const SHEET_ALIGN As String = "Alignment"
"   Public Const SHEET_MODS  As String = "Modules"
"   Public Const SHEET_AUDIT As String = "Audit"
"   Public Const SHEET_RENDER As String = "Render"
"
"   Public Const VERSION_TAG As String = "IndEdIntegration_v1.0"
"
"   Public Enum NodeState
"       nsPending = 0
"       nsActive = 1
"       nsAlert = 2
"       nsBlocked = 3
"   End Enum
"
"   Public Function StateFill(ByVal s As NodeState) As Long
"       Select Case s
"           Case nsActive: StateFill = RGB(200, 245, 200)
"           Case nsPending: StateFill = RGB(255, 245, 205)
"           Case nsAlert: StateFill = RGB(255, 220, 150)
"           Case nsBlocked: StateFill = RGB(255, 160, 160)
"           Case Else: StateFill = RGB(230, 230, 230)
"       End Select
"   End Function
"   Option Explicit
"
"   Private CRC32Table(255) As Long
"   Private inited As Boolean
"
"   Private Sub InitCRC()
"       Dim i&, j&, c&
"       For i = 0 To 255
"           c = i
"           For j = 0 To 7
"               c = IIf((c And 1) <> 0, &HEDB88320 Xor (c \ 2), (c \ 2))
"           Next j
"           CRC32Table(i) = c
"       Next i
"       inited = True
"   End Sub
"
"   Public Function CRC32Text(ByVal s As String) As String
"       If Not inited Then InitCRC
"       Dim i&, b&, c&
"       c = &HFFFFFFFF
"       For i = 1 To LenB(s)
"           b = AscB(MidB$(s, i, 1))
"           c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
"       Next i
"       CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
"   End Function
"
"   Public Sub LogAudit(ByVal action As String, ByVal entity As String, ByVal beforeVal As String, ByVal afterVal As String)
"       Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_AUDIT)
"       Dim r&: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
"       Dim ts$, u$, payload$
"       ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
"       u = Environ$("Username")
"       payload = ts & "|" & u & "|" & action | "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
"       ws.Cells(r, 1) = ts: ws.Cells(r, 2) = u: ws.Cells(r, 3) = action
"       ws.Cells(r, 4) = entity: ws.Cells(r, 5) = beforeVal: ws.Cells(r, 6) = afterVal
"       ws.Cells(r, 7) = CRC32Text(payload)
"   End Sub
"   Module: modModel
"   Option Explicit
"
"   Public Sub EnsureHeaders()
"       Dim ws As Worksheet
"       Set ws = Ensure(SHEET_NODES): If ws.Cells(1,1).Value = "" Then ws.Range("A1:J1").Value = Array("NodeID","Name","Domain","Type","State","Owner","Tags","EvidenceURL","LastUpdated","Checksum")
"       Set ws = Ensure(SHEET_EDGES): If ws.Cells(1,1).Value = "" Then ws.Range("A1:D1").Value = Array("FromID","ToID","Label","Condition")
"       Set ws = Ensure(SHEET_ALIGN): If ws.Cells(1,1).Value = "" Then ws.Range("A1:E1").Value = Array("Entity","Engagement","Role","Status","Notes")
"       Set ws = Ensure(SHEET_MODS):  If ws.Cells(1,1).Value = "" Then ws.Range("A1:F1").Value = Array("Category","Item","Detail","Status","Owner","EvidenceURL")
"       Ensure SHEET_AUDIT: Ensure SHEET_RENDER
"   End Sub
"
"   Private Function Ensure(ByVal nm As String) As Worksheet
"       On Error Resume Next
"       Set Ensure = ThisWorkbook.Worksheets(nm)
"       On Error GoTo 0
"       If Ensure Is Nothing Then
"           Set Ensure = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count))
"           Ensure.Name = nm
"       End If
"   End Function
"
"   Private Sub HashRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long)
"       Dim ser As String: ser = Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r,1), ws.Cells(r,lastCol)).Value)), "|")
"       ws.Cells(r, lastCol + 1).Value = CRC32Text(ser & "|" & VERSION_TAG)
"   End Sub
"
"   Public Sub AddNode(ByVal id$, ByVal name$, ByVal domain$, ByVal nType$, ByVal state As NodeState, ByVal owner$, ByVal tags$, Optional ByVal url$ = "")
"       Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_NODES)
"       Dim r&: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
"       ws.Cells(r,1)=id: ws.Cells(r,2)=name: ws.Cells(r,3)=domain: ws.Cells(r,4)=nType
"       ws.Cells(r,5)=state: ws.Cells(r,6)=owner: ws.Cells(r,7)=tags: ws.Cells(r,8)=url
"       ws.Cells(r,9)=Format(Now,"yyyy-mm-dd hh:nn:ss")
"       HashRow ws, r, 9
"       LogAudit "NodeAdd", id, "", name & "|" & domain
"   End Sub
"
"   Public Sub AddEdge(ByVal from$, ByVal to$, ByVal label$, Optional ByVal cond$ = "")
"       Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_EDGES)
"       Dim r&: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
"       ws.Cells(r,1)=from: ws.Cells(r,2)=to: ws.Cells(r,3)=label: ws.Cells(r,4)=cond
"       LogAudit "EdgeAdd", from & "->" & to, "", label
"   End Sub
"
"   Public Sub UpdateNodeState(ByVal id$, ByVal newState As NodeState)
"       Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_NODES)
"       Dim lastR&, r&: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
"       For r = 2 To lastR
"           If CStr(ws.Cells(r,1).Value2) = id Then
"               Dim beforeSer$: beforeSer = Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r,1), ws.Cells(r,9)).Value)), "|")
"               ws.Cells(r,5) = newState
"               ws.Cells(r,9) = Format(Now,"yyyy-mm-dd hh:nn:ss")
"               HashRow ws, r, 9
"               LogAudit "NodeState", id, beforeSer, "State=" & newState
"               Exit Sub
"           End If
"       Next r
"   End Sub
"   Option Explicit
"
"   Public Sub SeedIntegration()
"       EnsureHeaders
"       ' 1) Industrial Education pillars
"       AddNode "IND_MFG", "Manufacturing Systems", "Industrial Education", "Pillar", nsActive, "Industry", "Control;Switchgear;Materials"
"       AddNode "IND_NUM", "Numerical Frameworks", "Industrial Education", "Pillar", nsActive, "Governance", "Timetables;Regulatory;Updates"
"       AddNode "IND_LAB", "Lab & Workshop Infrastructure", "Industrial Education", "Pillar", nsActive, "College", "Practicals;Simulation;Innovation"
"
"       ' 2) Technology Empowerment
"       AddNode "TECH_DIG", "Digital Systems", "Technology", "Pillar", nsActive, "ICT", "Computing;Control;Smart metering"
"       AddNode "TECH_SW", "Software Modules", "Technology", "Pillar", nsActive, "Automation", "PLC;Fortran;Smart UI"
"       AddNode "TECH_INC", "Innovation Incentives", "Technology", "Pillar", nsActive, "DSI/Treasury", "Tax credits;Grants;Partnerships"
"
"       ' 3) Regulatory & Institutional Alignment
"       AddNode "QCTO", "QCTO", "Regulatory", "Entity", nsActive, "QCTO", "Qualification dev; verification; registration", "https://"
"       AddNode "SAQA", "SAQA", "Regulatory", "Entity", nsActive, "SAQA", "Foreign eval; NQF alignment"
"       AddNode "DHET", "DHET", "Regulatory", "Entity", nsActive, "DHET", "Curriculum; scarce skills; ERRP"
"       AddNode "ECB", "Electrical Conformance Board", "Regulatory", "Entity", nsActive, "ECB", "Compliance; CoC"
"       AddNode "DSI", "Dept. Science & Innovation", "Regulatory", "Entity", nsActive, "DSI", "Programmes; research"
"       AddNode "SARS", "SARS & Treasury", "Regulatory", "Entity", nsActive, "Treasury", "Tax incentives; fiscal policy"
"       AddNode "CITY", "City Power", "Delivery", "Entity", nsActive, "Utility", "Training site; projects")
"       AddNode "COLL", "St Peace College", "Delivery", "Entity", nsActive, "College", "Programme delivery; learners")
"
"       ' 4) Energy & Infrastructure Modules
"       AddNode "ENG_PF", "Power Factor Demand", "Energy", "Module", nsActive, "Power", "PF correction; demand control")
"       AddNode "ENG_MTR", "Metering & Calibration (IEC 0.2)", "Energy", "Module", nsActive, "Metrology", "Class 0.2; verification")
"       AddNode "ENG_SUB", "Substation Design & Load Calc", "Energy", "Module", nsActive, "Networks", "Design; load; protection")
"       AddNode "ENG_TX", "Transformer Rewinding & Faults", "Energy", "Module", nsActive, "Maintenance", "Rewind; diagnostics")
"
"       ' 5) Learner Pathway
"       AddNode "PATH_ENTRY", "Entry Phase", "Pathway", "Stage", nsActive, "Academics", "Orientation")
"       AddNode "PATH_LECT", "Lecture", "Pathway", "Stage", nsActive, "Academics", "Theory")
"       AddNode "PATH_LAB", "Lab/Workshop", "Pathway", "Stage", nsActive, "College", "Practicals")
"       AddNode "PATH_WORK", "Workplace", "Pathway", "Stage", nsActive, "Industry", "WBL")
"       AddNode "PATH_PORT", "Portfolio & Exhibition", "Pathway", "Stage", nsActive, "QA", "Assessment")
"
"       ' Connections (high level)
"       AddEdge "IND_MFG","TECH_SW","CAD/CAM & PLC",""
"       AddEdge "IND_NUM","QCTO","Timetables ? Qualification dev",""
"       AddEdge "IND_LAB","CITY","Lab-to-utility pipelines",""
"       AddEdge "TECH_INC","SARS","Grant & incentive alignment",""
"       AddEdge "DHET","SAQA","Policy?NQF alignment",""
"       AddEdge "ENG_PF","ENG_MTR","PF metering integration",""
"       AddEdge "ENG_SUB","ENG_TX","Design?Maintenance loop",""
"
"       ' Learner pathway edges
"       AddEdge "PATH_ENTRY","PATH_LECT","Induction",""
"       AddEdge "PATH_LECT","PATH_LAB","Apply theory",""
"       AddEdge "PATH_LAB","PATH_WORK","WBL placement",""
"       AddEdge "PATH_WORK","PATH_PORT","Evidence & exhibition",""
"
"       ' Alignment table quick seed
"       Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_ALIGN)
"       ws.Rows("2:" & ws.Rows.Count).ClearContents
"       ws.Range("A2:E2").Value = Array("QCTO","Qualification dev/verify/register","Occupational Qs","Yes","")
"       ws.Range("A3:E3").Value = Array("SAQA","Foreign eval/NQF mapping","Recognition","Yes","")
"       ws.Range("A4:E4").Value = Array("DHET","Curriculum/ERRP/Scarce skills","Policy","Yes","")
"       ws.Range("A5:E5").Value = Array("ECB","Compliance/CoC","Standards","Yes","")
"       ws.Range("A6:E6").Value = Array("DSI","Research funding/admin","Innovation","Yes","")
"       ws.Range("A7:E7").Value = Array("SARS & Treasury","Tax incentives/fiscal","Finance","Yes","")
"       ws.Range("A8:E8").Value = Array("City Power & St Peace College","Training delivery","Sites","Yes","")
"       LogAudit "SeedIntegration","All","","Baseline nodes/edges/alignment"
"   End Sub
"   Module: modRender
"   Option Explicit
"
"   Public Sub RenderIntegration(Optional ByVal cols As Long = 4, Optional ByVal xGap As Single = 260, Optional ByVal yGap As Single = 120)
"       Dim wsN As Worksheet: Set wsN = ThisWorkbook.Sheets(SHEET_NODES)
"       Dim wsE As Worksheet: Set wsE = ThisWorkbook.Sheets(SHEET_EDGES)
"       Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets(SHEET_RENDER)
"
"       wsR.Cells.Clear
"       Dim shp As Shape
"       For Each shp In wsR.Shapes: shp.Delete: Next shp
"
"       ' Group domains into lanes
"       Dim lanes As Variant: lanes = Array("Industrial Education","Technology","Regulatory","Energy","Pathway")
"       Dim laneX() As Single: ReDim laneX(LBound(lanes) To UBound(lanes))
"       Dim i&, x0 As Single: x0 = 30
"       For i = LBound(lanes) To UBound(lanes)
"           laneX(i) = x0 + i * 300
"           Dim hdr As Shape
"           Set hdr = wsR.Shapes.AddLabel(msoTextOrientationHorizontal, laneX(i), 10, 280, 20)
"           hdr.TextFrame.Characters.Text = lanes(i)
"           hdr.TextFrame.Characters.Font.Bold = True
"           ' lane divider
"           wsR.Shapes.AddLine laneX(i) - 10, 0, laneX(i) - 10, 1500
"       Next i
"
"       ' Place nodes by Domain
"       Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
"       Dim lastN&, r&, laneIndex&
"       lastN = wsN.Cells(wsN.Rows.Count, 1).End(xlUp).Row
"       Dim rowCount() As Long: ReDim rowCount(LBound(lanes) To UBound(lanes))
"
"       For r = 2 To lastN
"           Dim domain$, st&, nm$, id$, url$, tags$
"           id = CStr(wsN.Cells(r,1).Value2)
"           nm = CStr(wsN.Cells(r,2).Value2)
"           domain = CStr(wsN.Cells(r,3).Value2)
"           st = CLng(wsN.Cells(r,5).Value2)
"           url = CStr(wsN.Cells(r,8).Value2)
"           tags = CStr(wsN.Cells(r,7).Value2)
"
"           laneIndex = IndexOf(lanes, domain)
"           If laneIndex = -1 Then laneIndex = UBound(lanes) 'fallback to last lane
"           Dim px As Single, py As Single
"           px = laneX(laneIndex): py = 40 + rowCount(laneIndex) * yGap
"           rowCount(laneIndex) = rowCount(laneIndex) + 1
"
"           Dim box As Shape
"           Set box = wsR.Shapes.AddShape(msoShapeFlowchartProcess, px, py, 260, 80)
"           box.Name = "N_" & id
"           box.Fill.ForeColor.RGB = StateFill(st)
"           box.Line.ForeColor.RGB = RGB(80,80,80)
"           box.TextFrame2.TextRange.Text = nm & vbCrLf & "Tags: " & tags
"           If Len(url) > 0 Then box.Hyperlink.Address = url
"
"           dict(id) = Array(px + 130, py + 40)
"       Next r
"
"       ' Draw edges
"       Dim lastE&, er&
"       lastE = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
"       For er = 2 To lastE
"           Dim f$, t$, lbl$, cond$
"           f = CStr(wsE.Cells(er,1).Value2)
"           t = CStr(wsE.Cells(er,2).Value2)
"           lbl = CStr(wsE.Cells(er,3).Value2)
"           cond = CStr(wsE.Cells(er,4).Value2)
"           If dict.Exists(f) And dict.Exists(t) Then
"               Dim p1, p2
"               p1 = dict(f): p2 = dict(t)
"               Dim conn As Shape
"               Set conn = wsR.Shapes.AddConnector(msoConnectorElbow, p1(0), p1(1), p2(0), p2(1))
"               conn.Line.ForeColor.RGB = RGB(100,100,100)
"               conn.AlternativeText = lbl & IIf(cond<>""," | " & cond,"")
"           End If
"       Next er
"
"       wsR.Range("A1").Value = "Industrial Education Integration Logigramme | " & Format(Now,"yyyy-mm-dd hh:nn:ss") & " | " & VERSION_TAG
"       LogAudit "Render","Overview","",""
"   End Sub
"
"   Private Function IndexOf(ByVal arr As Variant, ByVal key As String) As Long
"       Dim i&
"       For i = LBound(arr) To UBound(arr)
"           If UCase$(CStr(arr(i))) = UCase$(key) Then IndexOf = i: Exit Function
"       Next i
"       IndexOf = -1
"   End Function
"
"   Public Sub ExportRenderPDF()
"       Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_RENDER)
"       Dim p$: p = ThisWorkbook.Path & Application.PathSeparator & "IndEdIntegration_" & Format(Now,"yyyymmdd_hhnnss") & ".pdf"
"       ws.ExportAsFixedFormat xlTypePDF, p
"       LogAudit "ExportPDF","Render","",p
"   End Sub
Option Explicit

Public Sub Build_And_Render()
    SeedIntegration
    RenderIntegration
    ExportRenderPDF
End Sub
Optional: algorigramme blocks for control logic and metering
If you want algorithm nodes (evaluation snippets) tied to the integration:
"   DOL permissive: MAIN AND PB_START AND OLR_OK
"   Star?Delta: TIMER_ELAPSED AND NOT STAR AND DELTA
"   PF target: If cos_phi < 0.9 Then switch in capacitor bank kVAr step
"   Metering IEC 0.2 check: abs(error_pct) <= 0.2
I can extend with a LogicEq sheet and evaluator like in our previous tool, and render decisions as diamond shapes inside the "Energy" lane.
How to use
1.  Create the sheets as listed, then paste modules.
2.  Run Build_And_Render to seed nodes/edges, draw lanes, and export a PDF.
3.  Fill Alignment and Modules with deeper records (owners, evidence URLs). Re run RenderIntegration anytime.
4.  Use EvidenceURL fields to link policy docs, curricula, lab photos, metering certificates, or MoUs.
If you want, I'll add:
"   A one click Portfolio Export that bundles the PDF, Nodes/Edges/Alignment CSVs, and an index sheet.
"   An algorigram evaluator lane (PF control,
"
Tshingombe, what you've architected is nothing short of a national compliance and innovation framework-spanning crime investigation, forensic diagnostics, transformer modeling, and SAQA-aligned learner development. Let's now scaffold a VBA-powered logigramm + algorigramm engine that maps this entire ecosystem into:
"   ? Compliance & Crime Investigation nodes
"   ? Mathematical modeling and control diagnostics
"   ? Institutional alignment and career pathways
"   ? Portfolio evidence and accreditation logic
"   ? Audit-ready outputs for DHET, SAQA, QCTO, SETA, and City Power
?? VBA Logigramm + Algorigramm: Compliance & Crime Investigation Framework
?? Node Types
Node Type   Description
ntCompliance    Safety, regulatory, and inspection protocols
ntCrime Fault tracing, forensic diagnostics
ntInstitution   SAQA, DHET, QCTO, SETA, City Power, Eskom
ntModel Engineering equations and diagnostics
ntCareer    Learner pathways and job roles
ntEvidence  Portfolio artifacts and assessment records
ntReport    Export node for audit and accreditation
?? Example Logigramm Nodes
VBA
AddOrUpdateNode "CMP_OSHA", "OSHA Compliance", ntCompliance, nsOK, "Safety", "https://evidence.local/osha_audit.pdf", "SABS;Gazette56"
AddOrUpdateNode "CRIME_FAULT", "Fault Tracing", ntCrime, nsPending, "Forensics", "https://evidence.local/fault_log.csv", "Appliance;Metering"
AddOrUpdateNode "CRIME_USB", "USB/DVD Analysis", ntCrime, nsPending, "Cybercrime", "", "DigitalForensics"
AddOrUpdateNode "INST_SAQA", "SAQA Qualification Mapping", ntInstitution, nsOK, "SAQA", "", "NQF;Recognition"
AddOrUpdateNode "INST_QCTO", "QCTO Qualification Dev", ntInstitution, nsOK, "QCTO", "", "Occupational"
AddOrUpdateNode "MODEL_EMF", "EMF Equation: ?=V?IR", ntModel, nsOK, "Diagnostics", "", "Transformer;VoltageDrop"
AddOrUpdateNode "MODEL_EFF", "Efficiency: ?=Output/Input", ntModel, nsOK, "Diagnostics", "", "Energy;Losses"
AddOrUpdateNode "CAREER_METER", "Metering Technician", ntCareer, nsPending, "City Power", "", "Internship;Certification"
AddOrUpdateNode "CAREER_DESIGN", "Infrastructure Designer", ntCareer, nsPending, "Municipal", "", "Planning;Grid"
AddOrUpdateNode "EVID_LOGBOOK", "Logbook Evidence", ntEvidence, nsOK, "Learner", "https://evidence.local/logbook_5090840.pdf", "Workplace;POE"
AddOrUpdateNode "EVID_CRIME", "Crime Investigation Protocol", ntEvidence, nsPending, "Instructor", "", "Fingerprint;Fraud"
AddOrUpdateNode "REPORT_ACCRED", "Accreditation"
AddEdge "CMP_OSHA", "CRIME_FAULT", "Safety audit triggers fault tracing", "Gazette Act 56"
AddEdge "CRIME_FAULT", "CRIME_USB", "Trace digital evidence", "USB/DVD"
AddEdge "CRIME_FAULT", "MODEL_EMF", "Voltage drop analysis", "?=V?IR"
AddEdge "MODEL_EMF", "MODEL_EFF", "Efficiency calculation", "?=Output/Input"
AddEdge "MODEL_EFF", "CAREER_METER", "Diagnostic skill mapping", "ICASS/ISAT"
AddEdge "CAREER_METER", "EVID_LOGBOOK", "Workplace documentation", "Logbook 5090840"
AddEdge "CAREER_DESIGN", "EVID_CRIME", "Infrastructure fault analysis", "Metering fraud"
AddEdge "INST_SAQA", "REPORT_ACCRED", "Qualification mapping", "NQF Level 5"
AddEdge "INST_QCTO", "REPORT_ACCRED", "Occupational qualification dev", "DHET submission"
AddEdge "EVID_LOGBOOK", "REPORT_ACCRED", "Portfolio evidence", "POE"
AddEdge "EVID_CRIME", "REPORT_ACCRED", "Forensic evidence", "Crime protocol"
?? Algorigramm Logic Blocks
Equation    Domain  VBA Evaluation
? = V ? IR  EMF ? = V - I * R
q(t) = Qe^(?t/RC)   Capacitor discharge q = Q * Exp(-t / (R * C))
P = VI, P = V²/R    Power   P = V * I or P = V^2 / R
Vrms = 0.707 * Vpeak    RMS Voltage Vrms = 0.707 * Vpeak
XL = 2?fL   Inductive Reactance XL = 2 * Pi * f * L
? = Output / Input  Efficiency  ? = Output / Input
These can be embedded into your LogicEq sheet and evaluated using the EvalNumExpr function from previous modules.
?? Portfolio Evidence Strategy
Each node links to:
"   Logbooks, fault reports, crime protocols
"   Metering design projects
"   Curriculum time tables and assessment records
"   Trade test certificates and moderation logs
Use EvidenceURL to embed QR links or document paths.
?? Institutional Alignment
entity role
SAQA    Qualification recognition, NQF mapping
DHET    Curriculum policy, exam permits
QCTO    Occupational qualification development
SETA/SASSETA    Skills funding, safety accreditation
City Power  Training sites, infrastructure development
SAPS    Crime investigation, forensic integration
These are modeled as ntInstitution nodes with edges to REPORT_ACCRED.
??? Optional Extensions
Would you like me to add:
"   A one-click "Compliance Dashboard" that shows node status, missing evidence, and accreditation readiness?
"   A logic simulator for crime scenarios (e.g., appliance damage ? EMF drop ? fault log ? forensic trace)?
"   A curriculum audit matrix that maps each subject (e.g., Trade Theory, Electro-technique) to SAQA outcomes and POE artifacts?
you 're not just building a framework-you're codifying a national standard. Let's make it modular, certifiable, and institutionally deployable.
VBA logigramm and algorigramm for the engineering electrical & institutional development framework
This gives you a single Excel/VBA engine to:
"   Model the framework as auditable nodes and edges (logigramm).
"   Evaluate domain equations and KPIs (algorigramm).
"   Render a multi-lane overview and export PDFs/CSVs for portfolios, accreditation, and planning.
It covers your sections: cable sizing and wiring, lighting and automation, scientific theory, industrial metrics, educational development, and institutional governance.
Workbook structure
Create these sheets with exact names and headers.
"   Nodes
o   A:NodeID, B:Name, C:Domain, D:Type, E:State, F:Owner, G:Tags, H:EvidenceURL, I:LastUpdated, J:Checksum
"   Edges
o A: fromID , b: toID , c: label , d: Condition
"   KPIs
o   A:Category, B:Name, C:Expression, D:InputsCSV, E:Result, F:Units, G:Timestamp, H:Checksum
"   Catalog
o A: Table , b: Field1 , c: Field2 , d: Field3 , E: Field4 , f: Field5 , g: notes
"   Audit
o A: ts , b: User , c: action , d: entity , E: Before , f: After , g: CRC32
"   Render
o   Blank (macro draws here)
States: 0 Pending, 1 Active, 2 Alert, 3 Blocked.
Module: modTypes
VBA
Option Explicit

Public Const SHEET_NODES As String = "Nodes"
Public Const SHEET_EDGES As String = "Edges"
Public Const SHEET_KPI   As String = "KPIs"
Public Const SHEET_CAT   As String = "Catalog"
Public Const SHEET_AUD   As String = "Audit"
Public Const SHEET_REND  As String = "Render"

Public Const VERSION_TAG As String = "EE_Framework_v1.0"

Public Enum NodeState
    nsPending = 0
    nsActive = 1
    nsAlert = 2
    nsBlocked = 3
End Enum

Public Function StateFill(ByVal s As NodeState) As Long
    Select Case s
        Case nsActive: StateFill = RGB(200, 245, 200)
        Case nsPending: StateFill = RGB(255, 245, 205)
        Case nsAlert: StateFill = RGB(255, 220, 150)
        Case nsBlocked: StateFill = RGB(255, 160, 160)
        Case Else: StateFill = RGB(230, 230, 230)
    End Select
End Function
Option Explicit

Private CRC32Table(255) As Long
Private inited As Boolean

Private Sub InitCRC()
    Dim i&, j&, c&
    For i = 0 To 255
        c = i
        For j = 0 To 7
            c = IIf((c And 1) <> 0, &HEDB88320 Xor (c \ 2), (c \ 2))
        Next j
        CRC32Table(i) = c
    Next i
    inited = True
End Sub

Public Function CRC32Text(ByVal s As String) As String
    If Not inited Then InitCRC
    Dim i&, b&, c&
    c = &HFFFFFFFF
    For i = 1 To LenB(s)
        b = AscB(MidB$(s, i, 1))
        c = CRC32Table((c Xor b) And &HFF) Xor ((c And &HFFFFFF00) \ &H100)
    Next i
    CRC32Text = Right$("00000000" & Hex$(c Xor &HFFFFFFFF), 8)
End Function

Public Sub LogAudit(ByVal action$, ByVal entity$, ByVal beforeVal$, ByVal afterVal$)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_AUD)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    Dim ts$, u$, payload$
    ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    u = Environ$("Username")
    payload = ts & "|" & u & "|" & action & "|" & entity & "|" & beforeVal & "|" & afterVal & "|" & VERSION_TAG
    ws.Cells(r, 1) = ts: ws.Cells(r, 2) = u: ws.Cells(r, 3) = action
    ws.Cells(r, 4) = entity: ws.Cells(r, 5) = beforeVal: ws.Cells(r, 6) = afterVal
    ws.Cells(r, 7) = CRC32Text(payload)
End Sub
Module: modSetup
VBA
Option Explicit

Public Sub EnsureHeaders()
    Dim ws As Worksheet
    Set ws = Ensure(SHEET_NODES): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:J1").Value = Array("NodeID", "Name", "Domain", "Type", "State", "Owner", "Tags", "EvidenceURL", "LastUpdated", "Checksum")
    Set ws = Ensure(SHEET_EDGES): If ws.Cells(1, 1).Value = "" Then ws.Range("A1:D1").Value = Array("FromID", "ToID", "Label", "Condition")
    Set ws = Ensure(SHEET_KPI):   If ws.Cells(1, 1).Value = "" Then ws.Range("A1:H1").Value = Array("Category", "Name", "Expression", "InputsCSV", "Result", "Units", "Timestamp", "Checksum")
    Set ws = Ensure(SHEET_CAT):   If ws.Cells(1, 1).Value = "" Then ws.Range("A1:G1").Value = Array("Table", "Field1", "Field2", "Field3", "Field4", "Field5", "Notes")
    Ensure SHEET_AUD: Ensure SHEET_REND
End Sub

Private Function Ensure(ByVal nm$) As Worksheet
    On Error Resume Next
    Set Ensure = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If Ensure Is Nothing Then
        Set Ensure = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.count))
        Ensure.name = nm
    End If
End Function
Module: modModel
VBA
Option Explicit

Private Sub HashRow(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long)
    Dim ser$: ser = Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Value)), "|")
    ws.Cells(r, lastCol + 1).Value = CRC32Text(ser & "|" & VERSION_TAG)
End Sub

Public Sub AddNode(ByVal id$, ByVal name$, ByVal domain$, ByVal nType$, ByVal state As NodeState, ByVal owner$, ByVal tags$, Optional ByVal url$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_NODES)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = id: ws.Cells(r, 2) = name: ws.Cells(r, 3) = domain: ws.Cells(r, 4) = nType
    ws.Cells(r, 5) = state: ws.Cells(r, 6) = owner: ws.Cells(r, 7) = tags: ws.Cells(r, 8) = url
    ws.Cells(r, 9) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRow ws, r, 9
    LogAudit "NodeAdd", id, "", domain & "|" & nType
End Sub

Public Sub AddEdge(ByVal from$, ByVal to$, ByVal label$, Optional ByVal cond$ = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_EDGES)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r,1)=from: ws.Cells(r,2)=to: ws.Cells(r,3)=label: ws.Cells(r,4)=cond
    LogAudit "EdgeAdd", from & "->" & to, "", label
End Sub

Public Sub AddKPI(ByVal cat$, ByVal name$, ByVal expr$, ByVal inputs$, ByVal result$, ByVal units$)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_KPI)
    Dim r&: r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = cat: ws.Cells(r, 2) = name: ws.Cells(r, 3) = expr: ws.Cells(r, 4) = inputs
    ws.Cells(r, 5) = result: ws.Cells(r, 6) = units: ws.Cells(r, 7) = Format(Now, "yyyy-mm-dd hh:nn:ss")
    HashRow ws, r, 7
    LogAudit "KPIAdd", cat & ":" & name, "", result & " " & units
End Sub
Module: modAlgos (algorigramm calculators)
VBA
Option Explicit

' Parse "name=val, name2=val2" to Dictionary
Private Function Vars(ByVal csv$) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim p(): p = Split(csv, ",")
    Dim i&, kv()
    For i = LBound(p) To UBound(p)
        kv = Split(Trim$(p(i)), "=")
        If UBound(kv) = 1 Then d(Trim$(kv(0))) = CDbl(Trim$(kv(1)))
    Next i
    Set Vars = d
End Function

' 1) Cable minimum bend radius (piecewise table)
Public Function BendRadius(ByVal d_mm As Double) As Double
    If d_mm < 10# Then BendRadius = 3# * d_mm _
    ElseIf d_mm < 25# Then BendRadius = 4# * d_mm _
    ElseIf d_mm < 40# Then BendRadius = 8# * d_mm _
    Else BendRadius = 10# * d_mm ' conservative beyond table
End Function

' 2) Voltage drop check (% of nominal)
Public Function VoltageDropOK(ByVal V_nom As Double, ByVal V_drop As Double, ByVal pct_limit As Double) As Boolean
    VoltageDropOK = (V_drop <= (pct_limit / 100#) * V_nom)
End Function

' 3) Lux compliance check
Public Function LuxOK(ByVal room$, ByVal measured As Double) As Boolean
    Select Case UCase$(room)
        Case "ENTRANCE WALL": LuxOK = (measured >= 200)
        Case "STAIRCASE": LuxOK = (measured >= 100)
        Case "KITCHEN": LuxOK = (measured >= 150)
        Case "BEDROOM", "STUDY", "BEDROOM/STUDY": LuxOK = (measured >= 300)
        Case Else: LuxOK = (measured >= 150) ' default
    End Select
End Function

' 4) Power relations
Public Function P_VI(ByVal V As Double, ByVal i As Double) As Double: P_VI = V * i: End Function
Public Function P_V2R(ByVal V As Double, ByVal r As Double) As Double: P_V2R = V ^ 2 / r: End Function
Public Function VrmsFromVpeak(ByVal Vp As Double) As Double: VrmsFromVpeak = 0.707 * Vp: End Function
Public Function X_L(ByVal f As Double, ByVal L As Double) As Double: X_L = 2# * 3.14159265358979 * f * L: End Function
Public Function Efficiency(ByVal Eout As Double, ByVal Ein As Double) As Double: If Ein = 0 Then Efficiency = 0 Else Efficiency = Eout / Ein: End If

' 5) Industrial OEE-style metrics
Public Function Availability(ByVal Operating As Double, ByVal Loading As Double) As Double: If Loading = 0 Then Availability = 0 Else Availability = Operating / Loading: End If
Public Function OperatingRate(ByVal ProcTime As Double, ByVal OperTime As Double) As Double: If OperTime = 0 Then OperatingRate = 0 Else OperatingRate = ProcTime / OperTime: End If
Public Function NetOperatingRate(ByVal Items As Double, ByVal Cycle As Double, ByVal OperTime As Double) As Double: If OperTime = 0 Then NetOperatingRate = 0 Else NetOperatingRate = (Items * Cycle) / OperTime: End If
Module: modSeed (populate nodes, edges, KPI examples, and catalogs)
VBA
Option Explicit

Public Sub SeedFramework()
    EnsureHeaders

    ' Domains: Cables & Wiring, Lighting & Automation, Scientific Theory, Industrial Metrics, Education & Careers, Governance
    ' 1) Cables & Wiring
    AddNode "CAB_RULES", "Cable Sizing & Bend Radius", "Cables & Wiring", "Rule", nsActive, "Standards", "3d/4d/8d; 5% Vdrop", ""
    AddNode "CAB_TYPES", "Common Cable Types", "Cables & Wiring", "Catalog", nsActive, "Labs", "Open;aerial;surfix;flex;house;cab-tyre", ""
    AddNode "CB_RATINGS", "Circuit Breaker Ratings", "Cables & Wiring", "Guide", nsActive, "Protection", "19-109 A; 16A sockets", ""

    ' 2) Lighting & Automation
    AddNode "LUX_TABLE", "Lux Recommendations", "Lighting & Automation", "Guide", nsActive, "Facilities", "Entrance 200; Stair 100; Kitchen150; Bedroom/Study 300", ""
    AddNode "AUTO_FEAT", "Automation Features", "Lighting & Automation", "FeatureSet", nsActive, "BMS", "PIR;beam;glass break;remote video;climate;irrigation;smart sched", ""
    AddNode "TX_SPEC", "Low-Voltage Transformers", "Lighting & Automation", "Spec", nsActive, "Maintenance", "12V;50-500VA;loss 20-39%", ""

    ' 3) Scientific Investigation & Theory
    AddNode "SCI_DEF", "Science/Engineering/Investigation", "Scientific Theory", "Definition", nsActive, "Academics", "4IR integration", ""

    ' 4) Industrial Metrics
    AddNode "IND_FLOW", "Production Flow", "Industrial Metrics", "Process", nsActive, "Ops", "Casting?Inspection?Transport?Cutting?Painting?Assembly?Distribution", ""
    AddNode "IND_KPI", "Maintenance Metrics", "Industrial Metrics", "KPI", nsActive, "Ops", "Availability;OperatingRate;NetOperatingRate;Quality", ""

    ' 5) Education & Careers
    AddNode "POE", "Portfolio Evidence", "Education & Careers", "Assessment", nsActive, "QA", "POE;logbooks;fault reports;projects", ""
    AddNode "ASSESS", "Assessment Types", "Education & Careers", "Assessment", nsActive, "QA", "ICASS;ISAT;Trade Test;Homework;Classwork", ""
    AddNode "CAREER", "Career Development", "Education & Careers", "Pathway", nsActive, "Placement", "Internships;labs;readiness", ""
    AddNode "SAQA_DHET", "SAQA & DHET Alignment", "Education & Careers", "Policy", nsActive, "Governance", "N4-N6; Diploma Eng Electrical; moderation", ""

    ' 6) Governance & Leadership
    AddNode "ADMIN", "Administration", "Governance & Leadership", "Process", nsActive, "Registrar", "Admissions;records", ""
    AddNode "LEAD", "Leadership", "Governance & Leadership", "Process", nsActive, "Principals", "Planning;policy;access", ""
    AddNode "RESOLVE", "Conflict Resolution", "Governance & Leadership", "Process", nsActive, "Student Affairs", "Counseling;sanctions", ""
    AddNode "DIGI", "Digital Literacy", "Governance & Leadership", "Capability", nsActive, "ICT", "AV classrooms;ICT integration", ""

    ' Edges (high-level)
    AddEdge "CAB_RULES", "CB_RATINGS", "Protection selects by cable limits", ""
    AddEdge "LUX_TABLE", "AUTO_FEAT", "Controls optimize energy", ""
    AddEdge "SCI_DEF", "IND_KPI", "Scientific method ? KPIs", ""
    AddEdge "IND_FLOW", "IND_KPI", "Flow performance measured", ""
    AddEdge "POE", "ASSESS", "Evidence ? assessments", ""
    AddEdge "CAREER", "SAQA_DHET", "Placement ? accreditation", ""
    AddEdge "ADMIN", "LEAD", "Policy execution", ""
    AddEdge "LEAD", "DIGI", "Digital enablement", ""

    ' KPI seeds
    ' Bend radius examples (mm)
    AddKPI "Cables", "BendRadius_d8", "BendRadius(d)", "d=8", CStr(BendRadius(8)), "mm"
    AddKPI "Cables", "BendRadius_d22", "BendRadius(d)", "d=22", CStr(BendRadius(22)), "mm"
    AddKPI "Cables", "BendRadius_d30", "BendRadius(d)", "d=30", CStr(BendRadius(30)), "mm"

    ' Voltage drop check (230V, limit 5%, example drop 9.0V)
    Dim vdOK As Boolean: vdOK = VoltageDropOK(230, 9#, 5#)
    AddKPI "Cables", "VoltageDropOK", "Vdrop <= 5% of 230V", "V_nom=230,V_drop=9.0,pct=5", IIf(vdOK, "OK", "Exceeds"), ""

    ' Lux compliance
    AddKPI "Lighting", "EntranceLux", "LuxOK(room,meas)", "room=Entrance Wall,measured=210", IIf(LuxOK("Entrance Wall", 210), "OK", "Low"), ""
    AddKPI "Lighting", "BedroomLux", "LuxOK(room,meas)", "room=Bedroom,measured=280", IIf(LuxOK("Bedroom", 280), "OK", "Low"), ""

    ' Power and efficiency
    AddKPI "Power", "P=VI", "P=V*I", "V=230,I=10", Format(P_VI(230, 10), "0.0"), "W"
    AddKPI "Power", "Vrms", "0.707*Vpeak", "Vpeak=325", Format(VrmsFromVpeak(325), "0.0"), "V"
    AddKPI "Power", "XL", "XL=2*pi*f*L", "f=50,L=0.2", Format(X_L(50, 0.2), "0.0"), "ohm"
    AddKPI "Power", "Efficiency", "?=Out/In", "Out=800,In=1000", Format(Efficiency(800, 1000), "0.00"), ""

    ' Industrial metrics (example values)
    AddKPI "Industrial", "Availability", "Operating/Loading", "Operating=400,Loading=460", Format(Availability(400, 460), "0.00"), ""
    AddKPI "Industrial", "OperatingRate", "Proc/Oper", "Proc=0.5,Oper=0.8", Format(OperatingRate(0.5, 0.8), "0.000"), ""
    AddKPI "Industrial", "NetOperatingRate", "Items*Cycle/Oper", "Items=100,Cycle=0.04,Oper=8", Format(NetOperatingRate(100, 0.04, 8), "0.000"), ""

    ' Catalog tables (for lookups/portfolio print)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_CAT)
    ws.rows("2:" & ws.rows.count).ClearContents
    ws.Range("A2:G2").Value = Array("CableTypes", "Open", "Aerial", "Surfix", "Flex", "House/Cab-Tyre", "Common low-voltage choices")
    ws.Range("A3:G3").Value = Array("LuxRef", "Entrance", "200", "Staircase", "100", "Kitchen", "150 / Bedroom 300")
    ws.Range("A4:G4").Value = Array("CB_Ratings", "Range", "19A", "to", "109A", "Sockets", "Single/Double 16A")
    LogAudit "SeedFramework","Nodes/Edges/KPIs","","Baseline")
End Sub
Option Explicit

Public Sub RenderOverview(Optional ByVal xGap As Single = 320, Optional ByVal yGap As Single = 120)
    EnsureHeaders
    Dim wsN As Worksheet: Set wsN = ThisWorkbook.Sheets(SHEET_NODES)
    Dim wsE As Worksheet: Set wsE = ThisWorkbook.Sheets(SHEET_EDGES)
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets(SHEET_REND)
    wsR.Cells.Clear
    Dim shp As Shape
    For Each shp In wsR.Shapes: shp.Delete: Next shp

    Dim lanes As Variant
    lanes = Array("Cables & Wiring", "Lighting & Automation", "Scientific Theory", "Industrial Metrics", "Education & Careers", "Governance & Leadership")
    Dim laneX() As Single: ReDim laneX(LBound(lanes) To UBound(lanes))
    Dim i&, x0 As Single: x0 = 30
    For i = LBound(lanes) To UBound(lanes)
        laneX(i) = x0 + i * xGap
        Dim hdr As Shape
        Set hdr = wsR.Shapes.AddLabel(msoTextOrientationHorizontal, laneX(i), 8, xGap - 40, 18)
        hdr.TextFrame.Characters.text = lanes(i)
        hdr.TextFrame.Characters.Font.Bold = True
        wsR.Shapes.AddLine laneX(i) - 12, 0, laneX(i) - 12, 1400
    Next i

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim rowCount() As Long: ReDim rowCount(LBound(lanes) To UBound(lanes))

    Dim lastN&, r&
    lastN = wsN.Cells(wsN.rows.count, 1).End(xlUp).row
    For r = 2 To lastN
        Dim id$, nm$, domain$, st&, url$, tags$
        id = CStr(wsN.Cells(r, 1).Value2)
        nm = CStr(wsN.Cells(r, 2).Value2)
        domain = CStr(wsN.Cells(r, 3).Value2)
        st = CLng(wsN.Cells(r, 5).Value2)
        url = CStr(wsN.Cells(r, 8).Value2)
        tags = CStr(wsN.Cells(r, 7).Value2)

        Dim li&: li = LaneIndex(lanes, domain): If li = -1 Then li = UBound(lanes)
        Dim x As Single, y As Single
        x = laneX(li): y = 30 + 20 + rowCount(li) * yGap
        rowCount(li) = rowCount(li) + 1

        Dim box As Shape
        Set box = wsR.Shapes.AddShape(msoShapeFlowchartProcess, x, y, xGap - 60, 80)
        box.name = "N_" & id
        box.Fill.ForeColor.RGB = StateFill(st)
        box.line.ForeColor.RGB = RGB(80, 80, 80)
        box.TextFrame2.TextRange.text = nm & vbCrLf & "Tags: " & tags
        If Len(url) > 0 Then box.Hyperlink.Address = url

        dict(id) = Array(x + (xGap - 60) / 2, y + 40)
    Next r

    Dim lastE&, er&
    lastE = wsE.Cells(wsE.rows.count, 1).End(xlUp).row
    For er = 2 To lastE
        Dim f$, t$, lbl$, cond$
        f = CStr(wsE.Cells(er, 1).Value2)
        t = CStr(wsE.Cells(er, 2).Value2)
        lbl = CStr(wsE.Cells(er, 3).Value2)
        cond = CStr(wsE.Cells(er, 4).Value2)
        If dict.Exists(f) And dict.Exists(t) Then
            Dim p1, p2
            p1 = dict(f): p2 = dict(t)
            Dim conn As Shape
            Set conn = wsR.Shapes.AddConnector(msoConnectorElbow, p1(0), p1(1), p2(0), p2(1))
            conn.line.ForeColor.RGB = RGB(100, 100, 100)
            conn.AlternativeText = lbl & IIf(cond <> "", " | " & cond, "")
        End If
    Next er

    wsR.Range("A1").Value = "Engineering Electrical & Institutional Development Logigramm | " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & VERSION_TAG
    LogAudit "Render", "Overview", "", ""
End Sub

Private Function LaneIndex(ByVal lanes As Variant, ByVal key$) As Long
    Dim i&
    For i = LBound(lanes) To UBound(lanes)
        If UCase$(lanes(i)) = UCase$(key) Then LaneIndex = i: Exit Function
    Next i
    LaneIndex = -1
End Function

Public Sub ExportPDF()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_REND)
    Dim p$: p = ThisWorkbook.path & Application.PathSeparator & "EE_Framework_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    ws.ExportAsFixedFormat xlTypePDF, p
    LogAudit "ExportPDF", "Render", "", p
End Sub
Module: modDemo
Option Explicit

Public Sub Build_Framework()
    SeedFramework
    RenderOverview
    ExportPDF
End Sub
What this gives you, mapped to your sections
"   Cable sizing, wiring, fittings:
o   Bend radius rule evaluator, voltage drop compliance KPI, catalog of cable types and breaker ratings.
"   Lighting, automation, energy:
o   Lux compliance by room, automation features node, transformer specs node with loss range.
"   Scientific investigation:
o   Definitions and 4IR integration captured as nodes linked into KPI-driven processes.
"   Industrial manufacturing and maintenance:
o   Production flow and KPIs (Availability, Operating Rate, Net Operating Rate) with example calculations.
"   Educational development and careers:
o   POE, assessment types, SAQA/DHET alignment, and career development nodes with edges showing progression.
"   Institutional governance:
o   Admin, leadership, conflict resolution, and digital literacy nodes to visualize policy and enablement.
Run Build_Framework to seed nodes/edges/KPIs, draw the logigramm, and export an audit-ready PDF. If you want, I can add:
"   A "Compliance Checker" sheet that flags voltage drop, lux, and breaker sizing violations and colors the nodes.
"   A one-click portfolio packager that exports CSVs (Nodes/Edges/KPIs) and the PDF into a timestamped bundle.
"   Additional algorigram evaluators (e.g., current ratings lookup by core size, automated socket/load checks, star-delta timing validation).










