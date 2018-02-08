# BC2

Provides a system tray icon to set and apply the Control Panel\Window colour
(hkcu\control panal\colors\window) via the right click menu. 

This key determines the background colour for Word documents, browser pages etc.
It is inteded for users who find a white background tiring on the eyes.

Unfortunates Windows 10 only honours the setting until the session is locked.  
When it is unlocked the page display reverts to White although the content of
the key hasn't changes. This appears to be a Windows 10 bug.

The program therefore detects session unlocked events and re-applies the required colour.
This requires separate threads for form display and event detection
