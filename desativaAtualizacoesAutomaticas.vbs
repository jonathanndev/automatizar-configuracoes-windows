Const AU_DISABLED = 1

Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
Set objSettings = objAutoUpdate.Settings

objSettings.NotificationLevel = AU_DISABLED
objSettings.Save