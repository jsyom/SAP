Option Explicit
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As GuiConnection
Public objSess As GuiSession
Public objSBar As GuiStatusbar
Public objSheet As Worksheet
Dim W_System
Dim iCtr As Integer
Const tcode = "BP"


Function Attach_Session(iRow, Optional mysystem As String) As Boolean
  Dim il, it
  Dim W_conn, W_Sess

  ' Unless a system is provided (XXXYYY where XXX is SID and YYY client)
  ' get the system from the sheet (in this case it is in cell A8)
  If mysystem = "" Then
      W_System = ActiveSheet.Cells(iRow, 1)
  Else
      W_System = mysystem
  End If
  ' If we are already connected to a session, exit do not try again
  If W_System = "" Then
     Attach_Session = False
     Exit Function
  End If
  ' If the session object is not nil, use that session (assume connected to the correct session)
  If Not objSess Is Nothing Then
      If objSess.Info.SystemName & objSess.Info.Client = W_System Then
          Attach_Session = True
          Exit Function
      End If
  End If
  ' If not connected to anything, set up the objects
  If objGui Is Nothing Then
     Set SapGuiAuto = GetObject("SAPGUI")
     Set objGui = SapGuiAuto.GetScriptingEngine
  End If
  ' Cycle through the open SAP GUI sessions and check which is in the same system running the matching transaction
  For il = 0 To objGui.Children.Count - 1
      Set W_conn = objGui.Children(il + 0)
      For it = 0 To W_conn.Children.Count - 1
          Set W_Sess = W_conn.Children(it + 0)
          If W_Sess.Info.SystemName & W_Sess.Info.Client = W_System And W_Sess.Info.Transaction = tcode Then
              Set objConn = objGui.Children(il + 0)
              Set objSess = objConn.Children(it + 0)
              Exit For
          End If
