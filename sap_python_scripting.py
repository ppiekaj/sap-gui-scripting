
import sys, win32com.client
import easygui
import pandas as pd
import time

def enter_item(session, item):
  try:

    #uruchomienie transakcji as01
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nas01"
    session.findById("wnd[0]").sendVKey(0)

    session.FindById("wnd[0]/usr/ctxtANLA-ANLKL").text = item[2]
    session.FindById("wnd[0]/usr/ctxtANLA-BUKRS").text = "GSG"
    session.FindById("wnd[0]/usr/txtRA02S-NASSETS").text = "1"
    session.FindById("wnd[0]/usr/ctxtRA02S-RANL1").setfocus()
    session.findById("wnd[0]").SendVKey(4)
    session.findById("wnd[1]/tbar[0]/btn[17]").Press()

    session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB007").select()
    session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB007/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[0,24]").text = item[1]
    session.findById("wnd[0]").SendVKey(0)
    session.findById("wnd[0]").SendVKey(0)
    session.findById("wnd[0]").SendVKey(0)
    session.findById("wnd[0]").SendVKey(0)

    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").Select()
    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1145/ctxtANLZ-KOSTL").text=str(item[3])

    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").Select()
    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1301/ctxtANLA-VMGLI").text = "b20"

    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB06").Select()
    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB06/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1401/ctxtANLV-VSART").text = "01"
    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB06/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1401/ctxtANLV-VSGES").text = "01"

    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB07").Select()
    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB07/ssubSUBSC:SAPLATAB:0200/subAREA5:SAPLXAIS:0990/ctxtANLU-ZZDYSPO").text = item[4]
    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB07/ssubSUBSC:SAPLATAB:0200/subAREA5:SAPLXAIS:0990/ctxtANLU-ZZKOMOR").text = str(item[5])
    session.FindById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB07/ssubSUBSC:SAPLATAB:0200/subAREA5:SAPLXAIS:0990/ctxtANLU-ZZRFIN2").text = str(item[6])

   
    #zapis 
    session.findById("wnd[0]/tbar[0]/btn[11]").Press()
    res = session.FindById("wnd[0]/sbar/pane[0]").text

  except:
    print(sys.exc_info())
    res = None
  finally:
    
    return res


def sap_connect():

  try:

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
      return None

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
      return None

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
      return None

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
      return None

  except:
    return None

  return session


def main():

  #podanie ścieżki
  path_ = easygui.fileopenbox("Wybierz plik z danymi", filetypes=["*.xlsx", "*.xls", "Excel files"])
  #wczytanie Excela
  df_items = pd.read_excel(path_)

  #do dorobienia obsługa błędu jeżeli nie wczytano pliku 
  #df_items = pd.DataFrame(df_items.iloc[0:3])
  df_items["Nr SAT"] = 0
  
  session = sap_connect()
  if session is None:
    easygui.msgbox(msg="Bład połączenia z SAP", title="Błąd")
    return
  for idx ,item in df_items.iterrows():
     result = enter_item(session, item )
     if result is None:
       result = "Wystąpił błąd"
     df_items.loc[idx,"Nr SAT"] = result

  df_items.to_excel("Wynik.xlsx", index=False)
    

if __name__ == "__main__":
    main()