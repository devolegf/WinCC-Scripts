Function CreateAlarm(AlarmId, text)
   Const hmiAlarmStateCome = 1
   Dim user, Alarm
   user = HMIRuntime.Tags("@CurrentUser").Read
   Set Alarm = HMIRuntime.Alarms(AlarmId)
   Alarm.ProcessValues(5).Value = "" & text
   Alarm.ProcessValues(4).Value = "" & user
   Alarm.State = hmiAlarmStateCome
   Alarm.UserName = user
   Alarm.Create
End Function
