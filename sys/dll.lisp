;;;; -*- Mode: LISP; Syntax: COMMON-LISP; Coding: shift_jis; -*-
(in-package :cl-win32ole-sys)

(eval-when (:compile-toplevel :load-toplevel :execute)
  (mapcan #'(lambda (lib)
              (cffi:load-foreign-library lib))
          '("ole32.dll"
            "oleaut32.dll"
            "user32.dll"
            "kernel32.dll"
            "advapi32.dll")))
