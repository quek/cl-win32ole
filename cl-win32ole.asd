;;;; -*- Mode: LISP; Syntax: COMMON-LISP; -*-
(asdf:defsystem :cl-win32ole
  :description "Win32 OLE high level Interface"
  :author "Yoshinori Tahara <read.eval.print@gmail.com>"
  :version "0.0.1"
  :licence "BSD"
  :components ((:module "api"
                        :serial t
                        :components
                        ((:file "package")
                         (:file "compat")
                         (:file "classes")
                         (:file "variant")
                         (:file "unknown")
                         (:file "dispatch")
                         (:file "type-info")
                         (:file "type-lib")
                         (:file "type-attr")
                         (:file "func-desc")
                         (:file "var-desc")
                         (:file "invoke")
                         )))
  :depends-on (cl-win32ole-sys))
