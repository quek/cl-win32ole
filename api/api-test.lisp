(eval-when (:compile-toplevel :load-toplevel :execute)
  (require :cl-win32ole))

(in-package :cl-win32ole)

;;(let ((dispatch (create-object "InternetExplorer.Application")))
(let ((dispatch (create-object "WScript.Shell")))
  (p (get-type-info-count dispatch))
  (let ((type-info (%get-type-info dispatch)))
    (p (get-name-of-documentation type-info))))
