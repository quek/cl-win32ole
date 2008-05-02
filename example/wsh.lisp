(eval-when (:compile-toplevel :load-toplevel :execute)
  (require :cl-win32ole)
  (use-package :cl-win32ole))

(defun wsh-example ()
  (let ((wsh (create-object "WScript.Shell")))
    (invoke wsh :popup "Hello")))

;;(wsh-example)
