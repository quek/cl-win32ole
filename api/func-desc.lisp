(in-package :cl-win32ole)

(defun make-func-desc (ptr-func-desc type-info)
  (let ((memid (cffi:foreign-slot-value ptr-func-desc 'FUNCDESC 'memid)))
    (make-instance
     'func-desc
     :name (get-name-of-documentation type-info memid)
     :memid memid
     :invkind (cffi:foreign-slot-value ptr-func-desc 'FUNCDESC 'invkind)
     )))
