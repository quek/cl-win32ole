(eval-when (:compile-toplevel :load-toplevel :execute)
  (asdf:oos 'asdf:load-op :cl-win32ole)
  (use-package :cl-win32ole))

(defun wsh-example ()
  (let ((wsh (create-object "WScript.Shell")))
    (invoke wsh :popup "Hello")))

;;(wsh-example)
