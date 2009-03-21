;;;; -*- Mode: LISP; Syntax: COMMON-LISP; -*-
(asdf:defsystem :cl-win32ole-sys
  :description "Win32 OLE low level Interface"
  :author "Yoshinori Tahara <read.eval.print@gmail.com>"
  :version "0.0.1"
  :licence "BSD"
  :components ((:module "sys"
                        :serial t
                        :components
                        ((:file "package")
                         (:file "util")
                         (:file "dll")
                         (:file "constant")
                         (:file "type")
                         (:file "iid")
                         (:file "ffi")
                         (:file "ole-variant")
                         (:file "ole")
                         (:file "ole-dispatch")
                         (:file "safearray")
                         (:file "systemtime")
                         )))
  :depends-on (cffi cl-ppcre trivial-garbage simple-date-time))
