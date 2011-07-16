(asdf:oos 'asdf:load-op :cl-win32ole)
(use-package :cl-win32ole)

#|
Dim aNodePath(0)
Set oServM = CreateObject("com.sun.star.ServiceManager")
Set oConfP = oServM.createInstance("com.sun.star.configuration.ConfigurationProvider")
Set aNodePath(0) = oServM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
aNodePath(0).Name = "nodepath"
aNodePath(0).Value = "/org.openoffice.Setup/Product"

Set oRegAccess = oConfP.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", aNodePath)

sOOVersion = oRegAccess.getByName("ooSetupVersionAboutBox")

MsgBox sOOVersion
|#

(let* ((sm (create-object "com.sun.star.ServiceManager"))
       (cp (ole sm :createInstance "com.sun.star.configuration.ConfigurationProvider"))
       (np (ole sm :Bridge_GetStruct "com.sun.star.beans.PropertyValue")))
  (with-slots (name value) np
    (setf name "nodepath")
    (setf value "/org.openoffice.Setup/Product")
    (let ((ra (ole cp :createInstanceWithArguments
		   "com.sun.star.configuration.ConfigurationAccess"
		   (list np))))
      (let ((v (ole ra :getByName "ooSetupVersionAboutBox")))
	(ole (create-object "WScript.Shell") :popup v)))))

#|
Dim aNoArgs()
set oServiceManager = CreateObject("com.sun.star.ServiceManager")
set oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
set oDoc = oDesktop.loadComponentFromURL("file:///C:/demo.odt", "_blank", 0, aNoArgs)
|#
(let* ((sm (create-object "com.sun.star.ServiceManager"))
       (dt (ole sm :createInstance "com.sun.star.frame.Desktop"))
       (doc (ole dt :loadComponentFromURL
		 "file:///C:/demo.odt"
		 "_blank"
		 0
		 (empty-array))))
  doc)
