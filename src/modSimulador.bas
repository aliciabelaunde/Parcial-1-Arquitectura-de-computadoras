'==============================================================================
'  SIMULADOR DE CPU DE 8 BITS  (arquitectura von Neumann / x86 simplificada)
'  Arquitectura de Computadoras SIS131 - UCB Santa Cruz - Parcial 1
'
'  Mapa de la hoja "Simulador":
'     A1:D7    Registros PC, IR, MAR, MDR, AX, BX (A2:B7), banderas ZF, CF, SF
'              (C2:D4) y registros CX, DX (C5:D6)
'     A9:D15   Panel de control (botones)
'     A17:D24  Unidad de control: fase, micro-operacion, ALU, reloj
'     A25:D25  Selector del programa demostrativo
'     A26:D30  Inspector de memoria (hex / binario / decimal / mnemonico)
'     F2:V18   RAM 256 x 8 bits  (datos en G3:V18)
'              filas 3-10 = segmento de CODIGO 00h-7Fh
'              filas 11-18 = segmento de DATOS 80h-FFh
'     Y2:AB20  Log de micro-operaciones (historial completo en hoja "Log")
'     AD1:AE12 Estado interno de la unidad de control (columnas ocultas)
'
'  Uso: ejecutar una vez ConfigurarSimulador (Alt+F8). Crea los botones,
'  los paneles y carga el programa demostrativo.
'==============================================================================
Option Explicit

' ---- Constantes de la hoja --------------------------------------------------
Private Const HOJA As String = "Simulador"
Private Const RAM_FILA As Long = 3        ' fila de la direccion 00h
Private Const RAM_COL As Long = 7         ' columna G = offset 0
Private Const LOG_COL As Long = 25        ' columna Y
Private Const LOG_FILA1 As Long = 3
Private Const LOG_FILAN As Long = 20
Private Const EST_COL As Long = 31        ' columna AE (valores de estado)

' ---- Registros de la CPU (se cargan desde la hoja en cada accion) -----------
Private PC As Long, IR As Long, IROP As Long, MAR As Long, MDR As Long
Private AX As Long, BX As Long, CX As Long, DX As Long, TMP As Long
Private ZF As Long, CF As Long, SF As Long

' ---- Estado de la unidad de control -----------------------------------------
Private Fase As String       ' fase que se ejecutara en el proximo pulso
Private Micro As Long        ' micro-operacion dentro de la fase
Private FaseUlt As String    ' fase de la ultima micro-operacion ejecutada
Private Ciclo As Long        ' numero de instruccion
Private Paso As Long         ' pulsos de reloj (micro-operaciones)
Private Halt As Boolean
Private Salto As Boolean
Private PCInstr As Long      ' direccion donde empieza la instruccion actual

' ---- Senales de la ultima micro-operacion (para resaltar y registrar) --------
Private Activos As String    ' ej. "|PC|MAR|"
Private MemAct As Long       ' celda de RAM accedida (-1 = ninguna)
Private MemModo As String    ' "R" lectura / "W" escritura
Private TextoRTL As String
Private Detalle As String
Private AluTxt As String

Private RAM(0 To 255) As Long
Private Pausa As Boolean
Private Ejecutando As Boolean

' ---- Instruccion decodificada -----------------------------------------------
Private Type TInstr
    Valida As Boolean
    Clase As String    ' NOP MOV LOAD STORE ALU UNARIA SALTO HLT
    Mnem As String
    Op1 As String      ' AX, BX, CX, DX, [DIR], [BX], DIR
    Op2 As String      ' AX, BX, CX, DX, IMM, [DIR], [BX]
    Tam As Long        ' 1 o 2 bytes
End Type

'==============================================================================
'  1. UTILIDADES
'==============================================================================
Private Function Sh() As Worksheet
    Set Sh = ThisWorkbook.Worksheets(HOJA)
End Function

Private Function B8(ByVal v As Long) As Long
    B8 = ((v Mod 256) + 256) Mod 256
End Function

Private Function H2(ByVal v As Long) As String
    H2 = Right$("0" & Hex$(B8(v)), 2)
End Function

Private Function HX(ByVal v As Long) As String
    HX = "0x" & H2(v)
End Function

Public Function Bin8(ByVal v As Long) As String
    Dim i As Long, s As String
    v = B8(v)
    For i = 7 To 0 Step -1
        If (v And (2 ^ i)) <> 0 Then s = s & "1" Else s = s & "0"
    Next i
    Bin8 = s
End Function

Private Function LeerHex(ByVal txt As Variant, ByRef ok As Boolean) As Long
    ' Convierte el texto de una celda ("3C", "0x3C", 7) a numero
    Dim s As String, i As Long, c As String, r As Long
    ok = False
    s = UCase$(Trim$(CStr(txt)))
    If Left$(s, 2) = "0X" Then s = Mid$(s, 3)
    If Right$(s, 1) = "H" Then s = Left$(s, Len(s) - 1)
    If s = "" Then LeerHex = 0: ok = True: Exit Function
    If Len(s) > 2 Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If InStr(1, "0123456789ABCDEF", c) = 0 Then Exit Function
        r = r * 16 + InStr(1, "0123456789ABCDEF", c) - 1
    Next i
    LeerHex = r
    ok = True
End Function

Private Function FL() As String
    FL = " " & ChrW(8592) & " "        ' flecha <-
End Function

Private Sub Activar(ByVal lista As String)
    Activos = Activos & Replace(lista, ",", "|") & "|"
End Sub

Private Function Activo(ByVal comp As String) As Boolean
    Activo = InStr(1, Activos, "|" & comp & "|") > 0
End Function

'==============================================================================
'  2. MEMORIA PRINCIPAL: operaciones primitivas Read / Write
'==============================================================================
Public Function MemRead(ByVal address As Long) As Long
    MemRead = RAM(B8(address))
End Function

Public Sub MemWrite(ByVal address As Long, ByVal value As Long)
    RAM(B8(address)) = B8(value)
    Sh.Cells(RAM_FILA + B8(address) \ 16, RAM_COL + B8(address) Mod 16).value = H2(value)
End Sub

'==============================================================================
'  3. DECODIFICADOR (ISA)
'     Codigo de registro (2 bits): 00 = AX, 01 = BX, 10 = CX, 11 = DX
'
'     00h        NOP
'     1dsh-7dsh  Operacion registro-registro. Nibble alto = operacion,
'                nibble bajo = dd ss (bits 3-2 destino, bits 1-0 fuente)
'                1 MOV  2 ADD  3 SUB  4 CMP  5 AND  6 OR  7 XOR
'                Ej.: 21h = 0010 0001 = ADD AX, BX ;  2Ah = ADD CX, CX
'     80h-9Bh    Operacion registro-inmediato: (op - 80h) \ 4 = operacion
'                (0 MOV 1 ADD 2 SUB 3 CMP 4 AND 5 OR 6 XOR), bits 1-0 = registro
'                Ej.: 80h = MOV AX, imm8 ; 8Ah = ADD CX, imm8
'     A0h-A3h LOAD reg,[dir8]   A4h-A7h STORE [dir8],reg
'     A8h-ABh LOAD reg,[BX]     ACh-AFh STORE [BX],reg   (indirecto)
'     B0h-B3h INC reg   B4h-B7h DEC reg   B8h-BBh NOT reg
'     C0h JMP  C1h JZ  C2h JNZ  C3h JC  C4h JNC  C5h JS  C6h JNS
'     FFh        HLT
'==============================================================================
Private Function RegCod(ByVal c As Long) As String
    RegCod = Choose((c And 3) + 1, "AX", "BX", "CX", "DX")
End Function

Private Function OpNombre(ByVal n As Long) As String
    OpNombre = Choose(n + 1, "MOV", "ADD", "SUB", "CMP", "AND", "OR", "XOR")
End Function

Private Function Decodificar(ByVal op As Long) As TInstr
    Dim d As TInstr
    d.Valida = True
    Select Case op
        Case &H0
            d.Clase = "NOP": d.Mnem = "NOP"
        Case &H10 To &H7F
            ' registro-registro: nibble alto = operacion, nibble bajo = dd ss
            d.Mnem = OpNombre((op \ 16) - 1)
            d.Op1 = RegCod((op \ 4) And 3)
            d.Op2 = RegCod(op And 3)
        Case &H80 To &H9B
            ' registro-inmediato
            d.Mnem = OpNombre((op - &H80) \ 4)
            d.Op1 = RegCod(op And 3)
            d.Op2 = "IMM"
        Case &HA0 To &HA3: d.Clase = "LOAD": d.Mnem = "LOAD": d.Op1 = RegCod(op): d.Op2 = "[DIR]"
        Case &HA4 To &HA7: d.Clase = "STORE": d.Mnem = "STORE": d.Op1 = "[DIR]": d.Op2 = RegCod(op)
        Case &HA8 To &HAB: d.Clase = "LOAD": d.Mnem = "LOAD": d.Op1 = RegCod(op): d.Op2 = "[BX]"
        Case &HAC To &HAF: d.Clase = "STORE": d.Mnem = "STORE": d.Op1 = "[BX]": d.Op2 = RegCod(op)
        Case &HB0 To &HBB
            d.Clase = "UNARIA"
            d.Mnem = Choose((op - &HB0) \ 4 + 1, "INC", "DEC", "NOT")
            d.Op1 = RegCod(op)
        Case &HC0 To &HC6
            d.Clase = "SALTO"
            d.Mnem = Choose(op - &HC0 + 1, "JMP", "JZ", "JNZ", "JC", "JNC", "JS", "JNS")
            d.Op1 = "DIR"
        Case &HFF
            d.Clase = "HLT": d.Mnem = "HLT"
        Case Else
            d.Valida = False: d.Clase = "INVALIDA": d.Mnem = "???"
    End Select
    If d.Clase = "" Then
        If d.Mnem = "MOV" Then d.Clase = "MOV" Else d.Clase = "ALU"
    End If
    If d.Op1 = "DIR" Or d.Op1 = "[DIR]" Or d.Op2 = "IMM" Or d.Op2 = "[DIR]" Then d.Tam = 2 Else d.Tam = 1
    Decodificar = d
End Function

Private Function TextoOp(ByVal s As String, ByVal operando As Long) As String
    Dim v As String
    If operando < 0 Then
        If s = "IMM" Then v = "imm8" Else v = "dir8"
    Else
        v = HX(operando)
    End If
    Select Case s
        Case "IMM", "DIR": TextoOp = v
        Case "[DIR]": TextoOp = "[" & v & "]"
        Case Else: TextoOp = s
    End Select
End Function

Private Function TextoInstr(ByRef d As TInstr, ByVal operando As Long) As String
    Dim s As String
    If Not d.Valida Then TextoInstr = "(opcode invalido)": Exit Function
    s = d.Mnem
    If d.Op1 <> "" Then s = s & " " & TextoOp(d.Op1, operando)
    If d.Op2 <> "" Then s = s & ", " & TextoOp(d.Op2, operando)
    TextoInstr = s
End Function

Private Function ModoDir(ByRef d As TInstr) As String
    If d.Op1 = "[DIR]" Or d.Op2 = "[DIR]" Then
        ModoDir = "directo"
    ElseIf d.Op1 = "[BX]" Or d.Op2 = "[BX]" Then
        ModoDir = "indirecto por registro"
    ElseIf d.Op2 = "IMM" Then
        ModoDir = "inmediato"
    ElseIf d.Op1 = "DIR" Then
        ModoDir = "absoluto (salto)"
    ElseIf d.Op1 <> "" Then
        ModoDir = "registro"
    Else
        ModoDir = "implicito"
    End If
End Function

Public Function Mnemonico(ByVal hexTxt As Variant) As String
    ' Funcion de hoja: =Mnemonico("22") devuelve "ADD AX, BX"
    Dim ok As Boolean, v As Long, d As TInstr
    v = LeerHex(hexTxt, ok)
    If Not ok Then Mnemonico = "?": Exit Function
    d = Decodificar(v)
    If d.Valida Then Mnemonico = TextoInstr(d, -1) Else Mnemonico = "dato " & v
End Function

'==============================================================================
'  4. ALU: operaciones y banderas (igual que x86)
'     ADD: CF=1 si hay acarreo (>255)   SUB/CMP: CF=1 si hay prestamo (a<b)
'     AND/OR/XOR: CF=0   INC/DEC: CF sin cambio   NOT: no modifica banderas
'     ZF=1 si el resultado es 0   SF = bit 7 del resultado
'==============================================================================
Private Function ALU(ByVal op As String, ByVal a As Long, ByVal b As Long) As Long
    Dim r As Long
    Select Case op
        Case "ADD": r = a + b: CF = IIf(r > 255, 1, 0)
        Case "SUB", "CMP": r = a - b: CF = IIf(a < b, 1, 0)
        Case "AND": r = a And b: CF = 0
        Case "OR": r = a Or b: CF = 0
        Case "XOR": r = a Xor b: CF = 0
        Case "INC": r = a + 1
        Case "DEC": r = a - 1
        Case "NOT"
            ALU = 255 - a
            AluTxt = "NOT " & HX(a) & " = " & HX(255 - a)
            Exit Function
    End Select
    r = B8(r)
    ZF = IIf(r = 0, 1, 0)
    SF = IIf((r And &H80) <> 0, 1, 0)
    ALU = r
    If op = "INC" Or op = "DEC" Then b = 1
    AluTxt = HX(a) & " " & op & " " & HX(b) & " = " & HX(r)
End Function

Private Function LeerReg(ByVal n As String) As Long
    Select Case n
        Case "AX": LeerReg = AX
        Case "BX": LeerReg = BX
        Case "CX": LeerReg = CX
        Case "DX": LeerReg = DX
    End Select
End Function

Private Sub EscribirReg(ByVal n As String, ByVal v As Long)
    Select Case n
        Case "AX": AX = B8(v)
        Case "BX": BX = B8(v)
        Case "CX": CX = B8(v)
        Case "DX": DX = B8(v)
    End Select
End Sub

Private Function Banderas() As String
    Banderas = "ZF=" & ZF & " CF=" & CF & " SF=" & SF
End Function

'==============================================================================
'  5. UNIDAD DE CONTROL: un pulso de reloj = una micro-operacion
'==============================================================================
Private Function NumMicro(ByVal f As String, ByRef d As TInstr) As Long
    Select Case f
        Case "FETCH": NumMicro = 3
        Case "DECODE": If d.Valida And d.Tam = 2 Then NumMicro = 4 Else NumMicro = 1
        Case "EXECUTE": If d.Clase = "LOAD" Or d.Clase = "STORE" Then NumMicro = 2 Else NumMicro = 1
        Case "STORE": NumMicro = 1
    End Select
End Function

Private Sub MicroPaso()
    Dim d As TInstr
    If Halt Then Exit Sub
    Activos = "|": MemAct = -1: MemModo = "": AluTxt = ""
    FaseUlt = Fase
    Paso = Paso + 1
    d = Decodificar(IR)
    Select Case Fase
        Case "FETCH": UFetch
        Case "DECODE": UDecode d
        Case "EXECUTE": UExecute d
        Case "STORE": UStore d
    End Select
    If Halt Then Exit Sub
    ' Secuenciador: avanza a la siguiente micro-operacion / fase
    d = Decodificar(IR)
    Micro = Micro + 1
    If Micro >= NumMicro(Fase, d) Then
        Micro = 0
        Select Case Fase
            Case "FETCH": Fase = "DECODE"
            Case "DECODE": Fase = "EXECUTE"
            Case "EXECUTE": Fase = "STORE"
            Case "STORE": Fase = "FETCH"
        End Select
    End If
End Sub

' ---- micro-operaciones de bus -------------------------------------------------
Private Sub UMar(ByVal origen As String, ByVal valor As Long)
    MAR = B8(valor)
    Activar origen & ",MAR,BUS"
    TextoRTL = "MAR" & FL() & origen
    Detalle = "MAR=" & HX(MAR)
End Sub

Private Sub ULeer()
    MDR = MemRead(MAR)
    MemAct = MAR: MemModo = "R"
    Activar "MAR,MDR,BUS"
    TextoRTL = "MDR" & FL() & "RAM[MAR]"
    Detalle = "RAM[" & HX(MAR) & "]=" & HX(MDR)
End Sub

' ---- FASE 1: FETCH --------------------------------------------------------
Private Sub UFetch()
    Dim d As TInstr
    Select Case Micro
        Case 0
            Ciclo = Ciclo + 1
            PCInstr = PC
            UMar "PC", PC
        Case 1
            ULeer
        Case 2
            IR = MDR: IROP = -1
            PC = B8(PC + 1)
            d = Decodificar(IR)
            Activar "MDR,IR,PC"
            TextoRTL = "IR" & FL() & "MDR ; PC" & FL() & "PC+1"
            Detalle = "IR=" & HX(IR) & " (" & d.Mnem & "), PC=" & HX(PC)
    End Select
End Sub

' ---- FASE 2: DECODE -------------------------------------------------------
Private Sub UDecode(ByRef d As TInstr)
    Select Case Micro
        Case 0
            Activar "IR,UC"
            If Not d.Valida Then
                Halt = True
                TextoRTL = "UC: opcode invalido, CPU detenida"
                Detalle = "IR=" & HX(IR) & " en " & HX(PCInstr)
                Exit Sub
            End If
            TextoRTL = "UC decodifica " & HX(IR) & " = " & TextoInstr(d, -1)
            Detalle = "modo " & ModoDir(d) & ", " & d.Tam & " byte(s)"
        Case 1
            UMar "PC", PC
            TextoRTL = TextoRTL & " (operando)"
        Case 2
            ULeer
        Case 3
            IROP = MDR
            PC = B8(PC + 1)
            Activar "MDR,IR,PC"
            TextoRTL = "IR(op)" & FL() & "MDR ; PC" & FL() & "PC+1"
            Detalle = "IR=" & H2(IR) & " " & H2(IROP) & " -> " & TextoInstr(d, IROP) & ", PC=" & HX(PC)
    End Select
End Sub

Private Function Fuente(ByRef d As TInstr, ByRef nombre As String) As Long
    If d.Op2 = "IMM" Then
        Fuente = IROP: nombre = "IR(op)"
    Else
        Fuente = LeerReg(d.Op2): nombre = d.Op2
    End If
End Function

' ---- FASE 3: EXECUTE ------------------------------------------------------
Private Sub UExecute(ByRef d As TInstr)
    Dim a As Long, b As Long, nb As String
    Select Case d.Clase
        Case "NOP"
            Activar "UC": TextoRTL = "NOP: sin operacion": Detalle = "-"
        Case "HLT"
            Activar "UC": TextoRTL = "UC: HLT, se prepara la parada del reloj": Detalle = "-"
        Case "MOV"
            TMP = Fuente(d, nb)
            Activar Replace(nb, "IR(op)", "IR") & ",TMP"
            TextoRTL = "TMP" & FL() & nb
            Detalle = "TMP=" & HX(TMP)
        Case "LOAD"
            If Micro = 0 Then
                If d.Op2 = "[BX]" Then
                    UMar "BX", BX                 ' direccionamiento indirecto
                Else
                    UMar "IR(op)", IROP
                    Activar "IR"
                End If
            Else
                ULeer
            End If
        Case "STORE"
            If Micro = 0 Then
                If d.Op1 = "[BX]" Then
                    UMar "BX", BX                 ' direccionamiento indirecto
                Else
                    UMar "IR(op)", IROP
                    Activar "IR"
                End If
            Else
                MDR = LeerReg(d.Op2)
                Activar d.Op2 & ",MDR"
                TextoRTL = "MDR" & FL() & d.Op2
                Detalle = "MDR=" & HX(MDR)
            End If
        Case "ALU", "UNARIA"
            a = LeerReg(d.Op1)
            If d.Clase = "ALU" Then b = Fuente(d, nb) Else nb = "1"
            TMP = ALU(d.Mnem, a, b)
            Activar d.Op1 & ",ALU,TMP,FLAGS"
            If d.Clase = "ALU" Then Activar Replace(nb, "IR(op)", "IR")
            If d.Mnem = "CMP" Then
                TextoRTL = "ALU: " & d.Op1 & " - " & nb & " (comparar)"
            ElseIf d.Mnem = "NOT" Then
                TextoRTL = "TMP" & FL() & "NOT " & d.Op1
            Else
                TextoRTL = "TMP" & FL() & d.Op1 & " " & d.Mnem & " " & nb
            End If
            Detalle = AluTxt & " | " & Banderas()
        Case "SALTO"
            Select Case d.Mnem
                Case "JMP": Salto = True
                Case "JZ": Salto = (ZF = 1)
                Case "JNZ": Salto = (ZF = 0)
                Case "JC": Salto = (CF = 1)
                Case "JNC": Salto = (CF = 0)
            End Select
            Activar "UC,FLAGS"
            TextoRTL = "UC evalua la condicion de " & d.Mnem
            Detalle = Banderas() & IIf(Salto, " -> salto TOMADO", " -> salto NO tomado")
    End Select
End Sub

' ---- FASE 4: STORE / WRITE-BACK --------------------------------------------
Private Sub UStore(ByRef d As TInstr)
    Select Case d.Clase
        Case "NOP"
            TextoRTL = "Sin write-back": Detalle = "-"
        Case "HLT"
            Halt = True
            Activar "UC"
            TextoRTL = "HLT: reloj detenido"
            Detalle = "Fin del programa (" & Ciclo & " instrucciones, " & Paso & " pulsos)"
        Case "MOV", "UNARIA"
            EscribirReg d.Op1, TMP
            Activar "TMP," & d.Op1
            TextoRTL = d.Op1 & FL() & "TMP"
            Detalle = d.Op1 & "=" & HX(TMP)
        Case "ALU"
            If d.Mnem = "CMP" Then
                Activar "FLAGS"
                TextoRTL = "Sin write-back: CMP solo cambia banderas"
                Detalle = Banderas()
            Else
                EscribirReg d.Op1, TMP
                Activar "TMP," & d.Op1
                TextoRTL = d.Op1 & FL() & "TMP"
                Detalle = d.Op1 & "=" & HX(TMP)
            End If
        Case "LOAD"
            EscribirReg d.Op1, MDR
            Activar "MDR," & d.Op1
            TextoRTL = d.Op1 & FL() & "MDR"
            Detalle = d.Op1 & "=" & HX(MDR)
        Case "STORE"
            MemWrite MAR, MDR
            MemAct = MAR: MemModo = "W"
            Activar "MAR,MDR,BUS"
            TextoRTL = "RAM[MAR]" & FL() & "MDR"
            Detalle = "RAM[" & HX(MAR) & "]=" & HX(MDR)
        Case "SALTO"
            If Salto Then
                PC = IROP
                Activar "IR,PC"
                TextoRTL = "PC" & FL() & "IR(op)"
                Detalle = "PC=" & HX(PC)
            Else
                Activar "PC"
                TextoRTL = "PC sin cambio"
                Detalle = "PC=" & HX(PC)
            End If
    End Select
End Sub

'==============================================================================
'  6. ESTADO: la hoja es la fuente de verdad (permite editar RAM en vivo)
'==============================================================================
Private Function CargarEstado() As Boolean
    Dim ws As Worksheet, v As Variant, r As Long, c As Long, ok As Boolean
    Dim partes() As String, t As String
    Set ws = Sh
    v = ws.Range(ws.Cells(RAM_FILA, RAM_COL), ws.Cells(RAM_FILA + 15, RAM_COL + 15)).value
    For r = 1 To 16
        For c = 1 To 16
            RAM((r - 1) * 16 + c - 1) = LeerHex(v(r, c), ok)
            If Not ok Then
                MsgBox "Valor invalido en la RAM, direccion " & H2((r - 1) * 16 + c - 1) & "h: '" & v(r, c) & "'" & vbLf & _
                       "Escriba un byte hexadecimal (00 a FF).", vbExclamation
                Exit Function
            End If
        Next c
    Next r
    PC = LeerHex(ws.Range("B2").value, ok)
    t = Trim$(CStr(ws.Range("B3").value))
    partes = Split(t & " ", " ")
    IR = LeerHex(partes(0), ok)
    If Trim$(partes(1)) = "" Or Trim$(partes(1)) = "--" Then IROP = -1 Else IROP = LeerHex(partes(1), ok)
    MAR = LeerHex(ws.Range("B4").value, ok)
    MDR = LeerHex(ws.Range("B5").value, ok)
    AX = LeerHex(ws.Range("B6").value, ok)
    BX = LeerHex(ws.Range("B7").value, ok)
    CX = LeerHex(ws.Range("D5").value, ok)
    DX = LeerHex(ws.Range("D6").value, ok)
    ZF = Val(ws.Range("D2").value): CF = Val(ws.Range("D3").value): SF = Val(ws.Range("D4").value)
    Fase = CStr(ws.Cells(1, EST_COL).value): If Fase = "" Then Fase = "FETCH"
    Micro = Val(ws.Cells(2, EST_COL).value)
    Ciclo = Val(ws.Cells(3, EST_COL).value)
    Paso = Val(ws.Cells(4, EST_COL).value)
    Halt = (ws.Cells(5, EST_COL).value = 1)
    Salto = (ws.Cells(6, EST_COL).value = 1)
    TMP = Val(ws.Cells(7, EST_COL).value)
    PCInstr = Val(ws.Cells(8, EST_COL).value)
    CargarEstado = True
End Function

Private Sub GuardarEstado()
    Dim ws As Worksheet
    Set ws = Sh
    ws.Range("B2").value = H2(PC)
    If IROP >= 0 Then ws.Range("B3").value = H2(IR) & " " & H2(IROP) Else ws.Range("B3").value = H2(IR)
    ws.Range("B4").value = H2(MAR)
    ws.Range("B5").value = H2(MDR)
    ws.Range("B6").value = H2(AX)
    ws.Range("B7").value = H2(BX)
    ws.Range("D5").value = H2(CX)
    ws.Range("D6").value = H2(DX)
    ws.Range("D2").value = ZF: ws.Range("D3").value = CF: ws.Range("D4").value = SF
    ws.Cells(1, EST_COL).value = Fase
    ws.Cells(2, EST_COL).value = Micro
    ws.Cells(3, EST_COL).value = Ciclo
    ws.Cells(4, EST_COL).value = Paso
    ws.Cells(5, EST_COL).value = IIf(Halt, 1, 0)
    ws.Cells(6, EST_COL).value = IIf(Salto, 1, 0)
    ws.Cells(7, EST_COL).value = TMP
    ws.Cells(8, EST_COL).value = PCInstr
End Sub

'==============================================================================
'  7. INTERFAZ: resaltado, paneles y log
'==============================================================================
Private Sub Pintar()
    Dim ws As Worksheet, i As Long, d As TInstr, nombres As Variant, col As Long, fases As Variant
    Set ws = Sh
    GuardarEstado

    ' Registros: resaltar los que participan en la micro-operacion
    nombres = Array("PC", "IR", "MAR", "MDR", "AX", "BX")
    For i = 0 To 5
        If Activo(CStr(nombres(i))) Then
            ws.Range("A" & (i + 2) & ":B" & (i + 2)).Interior.Color = RGB(255, 217, 102)
        ElseIf i Mod 2 = 0 Then
            ws.Range("A" & (i + 2) & ":B" & (i + 2)).Interior.Color = RGB(242, 246, 252)
        Else
            ws.Range("A" & (i + 2) & ":B" & (i + 2)).Interior.Color = RGB(255, 255, 255)
        End If
    Next i
    If Activo("FLAGS") Then
        ws.Range("C2:D4").Interior.Color = RGB(255, 217, 102)
    Else
        ws.Range("C2:D4").Interior.Color = RGB(255, 242, 204)
    End If
    nombres = Array("CX", "DX")
    For i = 0 To 1
        If Activo(CStr(nombres(i))) Then
            ws.Range("C" & (i + 5) & ":D" & (i + 5)).Interior.Color = RGB(255, 217, 102)
        Else
            ws.Range("C" & (i + 5) & ":D" & (i + 5)).Interior.Color = RGB(242, 246, 252)
        End If
    Next i

    ' RAM: colores de segmento + instruccion actual + celda accedida
    ws.Range(ws.Cells(3, RAM_COL), ws.Cells(10, RAM_COL + 15)).Interior.Color = RGB(221, 235, 247)
    ws.Range(ws.Cells(11, RAM_COL), ws.Cells(18, RAM_COL + 15)).Interior.Color = RGB(226, 239, 218)
    If Ciclo > 0 Then
        d = Decodificar(MemRead(PCInstr))
        For i = 0 To d.Tam - 1
            CeldaRAM(PCInstr + i).Interior.Color = RGB(255, 235, 156)
        Next i
    End If
    If MemAct >= 0 Then
        If MemModo = "W" Then
            CeldaRAM(MemAct).Interior.Color = RGB(255, 124, 128)
        Else
            CeldaRAM(MemAct).Interior.Color = RGB(146, 208, 80)
        End If
    End If

    ' Unidad de control: fase activa
    fases = Array("FETCH", "DECODE", "EXECUTE", "STORE")
    For i = 0 To 3
        col = i + 1
        If fases(i) = FaseUlt Then
            ws.Cells(18, col).Interior.Color = Choose(i + 1, RGB(47, 117, 181), RGB(112, 48, 160), RGB(237, 125, 49), RGB(84, 130, 53))
            ws.Cells(18, col).Font.Color = RGB(255, 255, 255)
        Else
            ws.Cells(18, col).Interior.Color = RGB(217, 217, 217)
            ws.Cells(18, col).Font.Color = RGB(89, 89, 89)
        End If
    Next i
    ws.Range("A19").value = TextoRTL
    ws.Range("A20").value = Detalle
    If Fase = "FETCH" And Micro > 0 Then
        ws.Range("B21").value = "buscando instruccion en " & HX(PCInstr) & "..."
    ElseIf Ciclo > 0 Then
        d = Decodificar(IR)
        ws.Range("B21").value = HX(PCInstr) & ": " & TextoInstr(d, IROP)
    Else
        ws.Range("B21").value = "-"
    End If
    If Activo("ALU") Then ws.Range("B22").value = AluTxt
    ws.Range("B22").Interior.Color = IIf(Activo("ALU"), RGB(255, 217, 102), RGB(255, 255, 255))
    ws.Range("B23").value = "Instr " & Ciclo & " | Pulso " & Paso & IIf(Halt, " | DETENIDA", " | Sig: " & Fase)
End Sub

Private Function CeldaRAM(ByVal addr As Long) As Range
    Set CeldaRAM = Sh.Cells(RAM_FILA + B8(addr) \ 16, RAM_COL + B8(addr) Mod 16)
End Function

Private Sub RegistrarLog()
    Dim ws As Worksheet, wl As Worksheet, linea As String, n As Long, r As Long, v As Variant
    Set ws = Sh
    linea = "[Paso " & Format$(Paso, "000") & "] " & FaseUlt & ": " & TextoRTL & " | " & Detalle
    ' Historial completo en la hoja Log
    Set wl = HojaLog()
    n = Paso + 1
    wl.Cells(n, 1).value = Paso
    wl.Cells(n, 2).value = Ciclo
    wl.Cells(n, 3).value = FaseUlt
    wl.Cells(n, 4).value = TextoRTL
    wl.Cells(n, 5).value = Detalle
    ' Consola: ultimas lineas (la mas reciente abajo)
    v = ws.Range(ws.Cells(LOG_FILA1, LOG_COL), ws.Cells(LOG_FILAN, LOG_COL)).value
    For r = 1 To LOG_FILAN - LOG_FILA1 + 1
        If CStr(v(r, 1)) = "" Then Exit For
    Next r
    If r > LOG_FILAN - LOG_FILA1 + 1 Then
        For r = 1 To LOG_FILAN - LOG_FILA1
            v(r, 1) = v(r + 1, 1)
        Next r
        r = LOG_FILAN - LOG_FILA1 + 1
    End If
    v(r, 1) = linea
    ws.Range(ws.Cells(LOG_FILA1, LOG_COL), ws.Cells(LOG_FILAN, LOG_COL)).value = v
End Sub

Private Sub LogSistema(ByVal texto As String)
    Dim ws As Worksheet
    Set ws = Sh
    ws.Range(ws.Cells(LOG_FILA1, LOG_COL), ws.Cells(LOG_FILAN, LOG_COL + 3)).ClearContents
    ws.Cells(LOG_FILA1, LOG_COL).value = "[SYSTEM] " & texto
End Sub

Private Function HojaLog() As Worksheet
    On Error Resume Next
    Set HojaLog = ThisWorkbook.Worksheets("Log")
    On Error GoTo 0
    If HojaLog Is Nothing Then
        ThisWorkbook.Worksheets.Add After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set HojaLog = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        HojaLog.Name = "Log"
        Sh.Activate
    End If
    If HojaLog.Range("A1").value = "" Then
        HojaLog.Range("A1:E1").value = Array("Paso", "Instruccion", "Fase", "Micro-operacion", "Detalle")
        HojaLog.Range("A1:E1").Font.Bold = True
    End If
End Function

Private Sub Esperar(ByVal ms As Long)
    Dim t0 As Single
    t0 = Timer
    Do While (Timer - t0) * 1000 < ms And (Timer - t0) >= 0
        DoEvents
        If Pausa Then Exit Do
    Loop
End Sub

'==============================================================================
'  8. BOTONES
'==============================================================================
Public Sub BtnStep()
    ' STEP: ejecuta UNA micro-operacion (un pulso de reloj)
    If Ejecutando Then Exit Sub
    If Not CargarEstado() Then Exit Sub
    If Halt Then MsgBox "La CPU esta detenida (HLT). Presione RESET.", vbInformation: Exit Sub
    MicroPaso
    RegistrarLog
    Pintar
End Sub

Public Sub BtnStepInstr()
    ' STEP INSTR: ejecuta todas las micro-operaciones de una instruccion
    Dim n As Long
    If Ejecutando Then Exit Sub
    If Not CargarEstado() Then Exit Sub
    If Halt Then MsgBox "La CPU esta detenida (HLT). Presione RESET.", vbInformation: Exit Sub
    Do
        MicroPaso
        RegistrarLog
        n = n + 1
    Loop Until Halt Or (Fase = "FETCH" And Micro = 0) Or n > 20
    Pintar
End Sub

Public Sub BtnRun()
    ' RUN: ejecucion continua con retardo ajustable (celda B24, en ms)
    Dim retardo As Long
    If Ejecutando Then Exit Sub
    If Not CargarEstado() Then Exit Sub
    If Halt Then MsgBox "La CPU esta detenida (HLT). Presione RESET.", vbInformation: Exit Sub
    retardo = Val(Sh.Range("B24").value)
    If retardo < 0 Then retardo = 0
    Pausa = False
    Ejecutando = True
    On Error GoTo Fin
    Application.EnableCancelKey = xlErrorHandler     ' ESC tambien pausa
    Sh.Range("A15").value = "EJECUTANDO..."
    Do While Not Halt And Not Pausa
        MicroPaso
        RegistrarLog
        Pintar
        DoEvents
        If retardo > 0 Then Esperar retardo
        If Paso > 20000 Then Exit Do                    ' proteccion contra bucles infinitos
    Loop
Fin:
    Ejecutando = False
    Application.EnableCancelKey = xlInterrupt
    GuardarEstado
    If Halt Then
        Sh.Range("A15").value = "DETENIDA (HLT)"
    Else
        Sh.Range("A15").value = "EN PAUSA"
    End If
End Sub

Public Sub BtnPause()
    Pausa = True
End Sub

Public Sub BtnReset()
    ' RESET: registros, banderas y PC a cero. La memoria se conserva.
    Dim ws As Worksheet
    If Ejecutando Then Pausa = True: Exit Sub
    If Not CargarEstado() Then Exit Sub
    PC = 0: IR = 0: IROP = -1: MAR = 0: MDR = 0: AX = 0: BX = 0: CX = 0: DX = 0: TMP = 0
    ZF = 0: CF = 0: SF = 0
    Fase = "FETCH": Micro = 0: FaseUlt = "": Ciclo = 0: Paso = 0: Halt = False: Salto = False: PCInstr = 0
    Activos = "|": MemAct = -1
    TextoRTL = "CPU reiniciada": Detalle = "PC=0x00"
    Set ws = Sh
    ws.Range("B22").value = "-"
    ws.Range("A15").value = "LISTA"
    HojaLog.Range("A2:E30000").ClearContents
    Pintar
    LogSistema "CPU reiniciada (registros y PC = 00h)."
End Sub

Public Sub BtnLoadProgram()
    ' LOAD PROGRAM: carga en la RAM el programa elegido en B25
    Dim nombre As String, i As Long
    If Ejecutando Then Exit Sub
    For i = 0 To 255
        RAM(i) = 0
    Next i
    nombre = Trim$(CStr(Sh.Range("B25").value))
    Select Case LCase$(Left$(nombre, 4))
        Case "fibo": ProgFibonacci
        Case "fact": ProgFactorial
        Case "cuen": ProgCuentaRegresiva
        Case Else: nombre = "Multiplicacion": ProgMultiplicacion
    End Select
    EscribirRAMCompleta
    BtnReset
    LogSistema "Programa cargado: " & nombre & "."
End Sub

Private Sub Poner(ByVal dirIni As Long, ByVal bytes As Variant)
    ' Escribe una lista de bytes en la RAM a partir de dirIni
    Dim i As Long
    For i = 0 To UBound(bytes)
        RAM(dirIni + i) = bytes(i)
    Next i
End Sub

Private Sub ProgMultiplicacion()
    '-----------------------------------------------------------------
    ' Multiplicacion por sumas sucesivas: RES = A x B (7 x 3 = 21 = 15h)
    '   Dir  Bytes   Etiqueta Instruccion         Comentario
    '   00   A1 80   INICIO   LOAD BX, [A]        BX = A (multiplicando)
    '   02   A2 81            LOAD CX, [B]        CX = B (contador)
    '   04   80 00            MOV AX, 0           AX = 0 (acumulador)
    '   06   8E 00            CMP CX, 0           B = 0?
    '   08   C1 0E            JZ FIN              si B = 0 el producto es 0
    '   0A   21      BUCLE    ADD AX, BX          AX = AX + A
    '   0B   B6               DEC CX              CX = CX - 1 (actualiza ZF)
    '   0C   C2 0A            JNZ BUCLE           repetir mientras CX <> 0
    '   0E   A4 82   FIN      STORE [RES], AX     RES = AX
    '   10   FF               HLT
    '   80   07      A        DB 7                multiplicando
    '   81   03      B        DB 3                multiplicador
    '   82   00      RES      DB 0                resultado
    '-----------------------------------------------------------------
    Poner &H0, Array(&HA1, &H80, &HA2, &H81, &H80, &H0, &H8E, &H0, &HC1, &HE, &H21, &HB6, &HC2, &HA, &HA4, &H82, &HFF)
    Poner &H80, Array(&H7, &H3, &H0)
End Sub

Private Sub ProgFibonacci()
    '-----------------------------------------------------------------
    ' Serie de Fibonacci hasta desbordar 8 bits (CF = 1). Guarda cada termino en A0h...
    '   Dir  Bytes   Etiqueta Instruccion         Comentario
    '   00   80 00   INICIO   MOV AX, 0           AX = F(n-1) = 0
    '   02   82 01            MOV CX, 1           CX = F(n) = 1
    '   04   81 A0            MOV BX, 0xA0        BX = puntero al arreglo SERIE
    '   06   AE      BUCLE    STORE [BX], CX      SERIE[i] = F(n)  (indirecto)
    '   07   B1               INC BX              siguiente posicion
    '   08   1C               MOV DX, AX          DX = F(n-1)
    '   09   2E               ADD DX, CX          DX = F(n-1) + F(n)
    '   0A   C3 10            JC FIN              CF = 1: ya no cabe en 8 bits
    '   0C   12               MOV AX, CX          F(n-1) = F(n)
    '   0D   1B               MOV CX, DX          F(n) = F(n+1)
    '   0E   C0 06            JMP BUCLE
    '   10   FF      FIN      HLT
    '-----------------------------------------------------------------
    Poner &H0, Array(&H80, &H0, &H82, &H1, &H81, &HA0, &HAE, &HB1, &H1C, &H2E, &HC3, &H10, &H12, &H1B, &HC0, &H6, &HFF)
End Sub

Private Sub ProgFactorial()
    '-----------------------------------------------------------------
    ' Factorial por sumas sucesivas: RES = N! (5! = 120 = 78h)
    '   Dir  Bytes   Etiqueta Instruccion         Comentario
    '   00   A2 80   INICIO   LOAD CX, [N]        CX = N
    '   02   80 01            MOV AX, 1           AX = 1 (resultado parcial)
    '   04   8E 02   EXT      CMP CX, 2
    '   06   C3 12            JC FIN              si CX < 2 terminar
    '   08   14               MOV BX, AX          BX = valor a sumar
    '   09   1E               MOV DX, CX
    '   0A   B7               DEC DX              DX = CX - 1 sumas
    '   0B   21      INT      ADD AX, BX          AX = AX + BX
    '   0C   B7               DEC DX
    '   0D   C2 0B            JNZ INT             AX = AX x CX por sumas sucesivas
    '   0F   B6               DEC CX
    '   10   C0 04            JMP EXT
    '   12   A4 81   FIN      STORE [RES], AX     RES = N!
    '   14   FF               HLT
    '   80   05      N        DB 5                N
    '   81   00      RES      DB 0                resultado (5! = 120 = 78h)
    '-----------------------------------------------------------------
    Poner &H0, Array(&HA2, &H80, &H80, &H1, &H8E, &H2, &HC3, &H12, &H14, &H1E, &HB7, &H21, &HB7, &HC2, &HB, &HB6, &HC0, &H4, &HA4, &H81, &HFF)
    Poner &H80, Array(&H5, &H0)
End Sub

Private Sub ProgCuentaRegresiva()
    '-----------------------------------------------------------------
    ' Cuenta regresiva desde N; guarda solo los pares en A0h... (10,8,6,4,2,0)
    '   Dir  Bytes   Etiqueta Instruccion         Comentario
    '   00   A2 80   INICIO   LOAD CX, [N]        CX = N
    '   02   81 A0            MOV BX, 0xA0        BX = puntero al arreglo PARES
    '   04   1E      BUCLE    MOV DX, CX
    '   05   93 01            AND DX, 1           ZF = 1 si CX es par
    '   07   C2 0B            JNZ SIGUE           impar: no se guarda
    '   09   AE               STORE [BX], CX      guardado condicional
    '   0A   B1               INC BX
    '   0B   8E 00   SIGUE    CMP CX, 0
    '   0D   C1 12            JZ FIN              llego a 0
    '   0F   B6               DEC CX
    '   10   C0 04            JMP BUCLE
    '   12   FF      FIN      HLT
    '   80   0A      N        DB 10               valor inicial
    '-----------------------------------------------------------------
    Poner &H0, Array(&HA2, &H80, &H81, &HA0, &H1E, &H93, &H1, &HC2, &HB, &HAE, &HB1, &H8E, &H0, &HC1, &H12, &HB6, &HC0, &H4, &HFF)
    Poner &H80, Array(&HA)
End Sub

Private Sub EscribirRAMCompleta()
    Dim v() As Variant, r As Long, c As Long
    ReDim v(1 To 16, 1 To 16)
    For r = 1 To 16
        For c = 1 To 16
            v(r, c) = H2(RAM((r - 1) * 16 + c - 1))
        Next c
    Next r
    Sh.Range(Sh.Cells(RAM_FILA, RAM_COL), Sh.Cells(RAM_FILA + 15, RAM_COL + 15)).value = v
End Sub

'==============================================================================
'  9. CONFIGURACION (ejecutar una sola vez)
'==============================================================================
Public Sub ConfigurarSimulador()
    Dim ws As Worksheet, i As Long
    Set ws = Sh
    Application.ScreenUpdating = False

    ' Textos de encabezado en espanol
    ws.Range("A1").value = "REGISTRO": ws.Range("B1").value = "VALOR"
    ws.Range("C1").value = "BANDERA": ws.Range("D1").value = "VALOR"
    ws.Range("W3").value = "CODIGO 00h-7Fh"
    For i = 0 To 15                                  ' etiquetas de fila de la RAM
        ws.Cells(RAM_FILA + i, RAM_COL - 1).NumberFormat = "@"
        ws.Cells(RAM_FILA + i, RAM_COL - 1).value = Hex$(i) & "0"
    Next i

    ' Celdas de RAM y registros como texto (para que "07" no se convierta en 7)
    ws.Range(ws.Cells(RAM_FILA, RAM_COL), ws.Cells(RAM_FILA + 15, RAM_COL + 15)).NumberFormat = "@"
    ws.Range("B2:B7").NumberFormat = "@"

    ' Registros CX y DX (C5:D6)
    ws.Range("C5").value = "CX (Contador)"
    ws.Range("C6").value = "DX (Datos)"
    ws.Range("C5:C6").Font.Bold = True
    ws.Range("C5:C6").Font.Color = RGB(31, 56, 100)
    ws.Range("D5:D6").NumberFormat = "@"
    ws.Range("D5:D6").HorizontalAlignment = xlCenter
    ws.Range("D5:D6").Font.Bold = True
    ws.Range("D5").value = "00"
    ws.Range("D6").value = "00"

    ' Selector del programa demostrativo (A25:D25)
    ws.Range("A25").value = "Programa"
    ws.Range("A25").Font.Bold = True
    ws.Range("B25:D25").Merge
    ws.Range("B25").HorizontalAlignment = xlCenter
    ws.Range("B25").Interior.Color = RGB(255, 255, 204)
    If Trim$(CStr(ws.Range("B25").value)) = "" Then ws.Range("B25").value = "Multiplicacion"
    On Error Resume Next                              ' la lista desplegable es opcional
    ws.Range("B25").Validation.Delete
    ws.Range("B25").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="Multiplicacion,Fibonacci,Factorial,Cuenta regresiva"
    On Error GoTo 0

    ' Panel de unidad de control (A17:D24)
    With ws.Range("A17:D17")
        .Merge
        .value = "UNIDAD DE CONTROL"
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("A18:D18").value = Array("FETCH", "DECODE", "EXECUTE", "STORE")
    ws.Range("A18:D18").Font.Bold = True
    ws.Range("A18:D18").HorizontalAlignment = xlCenter
    ws.Range("A18:D18").Interior.Color = RGB(217, 217, 217)
    For i = 19 To 20
        ws.Range("A" & i & ":D" & i).Merge
        ws.Range("A" & i).HorizontalAlignment = xlCenter
    Next i
    ws.Range("A19").Font.Name = "Consolas": ws.Range("A19").Font.Bold = True: ws.Range("A19").Font.Size = 12
    ws.Range("A20").Font.Name = "Consolas": ws.Range("A20").Font.Size = 9
    ws.Range("A21").value = "Instruccion": ws.Range("A22").value = "ALU": ws.Range("A23").value = "Reloj"
    ws.Range("A24").value = "Retardo RUN (ms)"
    For i = 21 To 23
        ws.Range("B" & i & ":D" & i).Merge
        ws.Range("B" & i).Font.Name = "Consolas"
    Next i
    ws.Range("A21:A24").Font.Bold = True
    ws.Range("B24").value = 300
    ws.Range("B24").Interior.Color = RGB(255, 255, 204)
    ws.Range("A17:D24").Borders.LineStyle = xlContinuous
    ws.Range("A17:D24").Borders.Color = RGB(191, 191, 191)

    ' Inspector de memoria (A26:D30): escriba una direccion en B27
    With ws.Range("A26:D26")
        .Merge
        .value = "INSPECTOR DE MEMORIA"
        .Interior.Color = RGB(51, 63, 79)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("A27").value = "Direccion (hex)": ws.Range("B27").NumberFormat = "@": ws.Range("B27").value = "10"
    ws.Range("B27").Interior.Color = RGB(255, 255, 204)
    ws.Range("A28").value = "Hex / Dec": ws.Range("A29").value = "Binario": ws.Range("A30").value = "Mnemonico"
    ws.Range("B28").Formula = "=INDEX($G$3:$V$18,INT(HEX2DEC($B$27)/16)+1,MOD(HEX2DEC($B$27),16)+1)"
    ws.Range("C28").Formula = "=HEX2DEC(B28)"
    ws.Range("B29").Formula = "=HEX2BIN(B28,8)"
    ws.Range("B30").Formula = "=Mnemonico(B28)"
    ws.Range("B30:D30").Merge
    ws.Range("A27:A30").Font.Bold = True
    ws.Range("A26:D30").Borders.LineStyle = xlContinuous
    ws.Range("A26:D30").Borders.Color = RGB(191, 191, 191)

    ' Estado interno (columnas AD:AE ocultas)
    ws.Range("AD1:AD8").value = Application.Transpose(Array("FASE", "MICRO", "CICLO", "PASO", "HALT", "SALTO", "TMP", "PCINSTR"))
    ws.Columns("AD:AE").Hidden = True

    CrearBotones
    Application.ScreenUpdating = True
    BtnLoadProgram
    MsgBox "Simulador configurado. Use STEP, STEP INSTR o RUN.", vbInformation
End Sub

Private Sub CrearBotones()
    Dim ws As Worksheet, shp As Shape, i As Long, x As Double, y As Double, w As Double, h As Double
    Dim textos As Variant, macros As Variant, colores As Variant
    Set ws = Sh
    ' Borrar botones anteriores del panel A9:D15
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.Top >= ws.Range("A9").Top - 2 And shp.Top < ws.Range("A16").Top And shp.Left < ws.Range("E1").Left Then shp.Delete
    Next i
    textos = Array("STEP (micro-op)", "STEP INSTR", "RUN", "PAUSE", "RESET", "LOAD PROGRAM")
    macros = Array("BtnStep", "BtnStepInstr", "BtnRun", "BtnPause", "BtnReset", "BtnLoadProgram")
    colores = Array(RGB(47, 117, 181), RGB(31, 78, 121), RGB(84, 130, 53), RGB(191, 143, 0), RGB(192, 0, 0), RGB(112, 48, 160))
    w = (ws.Range("A10:D10").Width - 18) / 2
    h = (ws.Range("A10:A14").Height - 18) / 3
    For i = 0 To 5
        x = ws.Range("A10").Left + 6 + (i Mod 2) * (w + 6)
        y = ws.Range("A10").Top + 4 + (i \ 2) * (h + 5)
        Set shp = ws.Shapes.AddShape(5, x, y, w, h)       ' 5 = rectangulo redondeado
        shp.Name = "btn" & macros(i)
        shp.Fill.ForeColor.RGB = colores(i)
        shp.Line.Visible = msoFalse
        shp.OnAction = macros(i)
        With shp.TextFrame
            .Characters.Text = textos(i)
            .Characters.Font.Bold = True
            .Characters.Font.Size = 10
            .Characters.Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    Next i
    With ws.Range("A15:D15")
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .value = "LISTA"
    End With
End Sub