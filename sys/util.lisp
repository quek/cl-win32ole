(in-package :cl-win32ole-sys)

;; for debug
(defmacro p (&body body)
  `(progn ,@(mapcar #'(lambda (arg)
                        `(format t "~30a ; => ~a~%" ',arg ,arg))
                    body)))
