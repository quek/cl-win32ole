(in-package :cl-win32ole)

(defun finalize (object lambda)
  #+sbcl
  (sb-ext:finalize object lambda))
