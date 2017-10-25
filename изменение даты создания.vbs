
  Set app = CreateObject("Shell.Application")
  Set folder = app.NameSpace("C:\ажЬЯ")
  Set file = folder.ParseName("1150000002499==16022015.txt")
file.ModifyDate = "2015-03-04"
