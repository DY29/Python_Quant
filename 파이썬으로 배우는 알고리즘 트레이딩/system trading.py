import win32com.client
instCpCybos = win32.com.client.Dispatch("CpUtil.CpCybos")
print(instCpCybos.IsConnect)
