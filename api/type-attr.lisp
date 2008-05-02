(in-package :cl-win32ole)

(defun make-type-attr (ptr-type-attr)
  (make-instance
   'type-attr
   :funcs (cffi:foreign-slot-value ptr-type-attr 'TYPEATTR 'cFuncs)
   :vars (cffi:foreign-slot-value ptr-type-attr 'TYPEATTR 'cVars)
   :imple-types (cffi:foreign-slot-value ptr-type-attr 'TYPEATTR 'cImplTypes)
   ))
