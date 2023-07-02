# TaskbarMeetingNotifier
Windows taskbar widget which shows upcoming meetings

Currently under development, already works in principle, altough recurring appointments are not shown yet.

![grafik](https://github.com/nilshellerhoff/TaskbarMeetingNotifier/assets/24147614/0316adaa-1f1d-4743-b46e-06fe7f10f89e)

Built using [Sharpshell Deskband](https://github.com/dwmkerr/sharpshell/blob/main/docs/extensions/deskband/deskband.md)

# Installing

- compile the project
- locate `RegAsm.exe` (is under `C:\Windows\Microsoft.NET\Framework64\<some version>\RegAsm.exe`)
- run the following from the project directory:
```
C:\Windows\Microsoft.NET\Framework64\<some version>\RegAsm.exe /codebase TaskbarMeetingNotifier\bin\Release\TaskbarMeetingNotifier.dll
```
- restart `explorer.exe`
- enable `TaskBar Meeting Notifier` from the taskbar context menu
![grafik](https://github.com/nilshellerhoff/TaskbarMeetingNotifier/assets/24147614/d1da6073-16e4-4801-a113-f51b4e2658db)

# TODOs / Ideas

- [ ] include recurring appointments
- [ ] make the UI nicer
- [ ] include a join option for online meetings
