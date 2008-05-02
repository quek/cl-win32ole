(in-package :cl-win32ole)

(defun make-type-lib (ptr)
  (make-instance 'type-lib :ptr ptr))

(defmethod get-type-info-count ((type-lib type-lib))
  (type-lib-get-type-info-count (ptr type-lib)))

(defmethod get-name-of-documentation ((type-lib type-lib) &optional (index -1))
  (type-lib-get-name-of-documentation (ptr type-lib) index))

(defmethod get-type-info ((type-lib type-lib) &optional index)
  (make-type-info (type-lib-get-type-info (ptr type-lib) index)))

(defmethod get-var-desc-list ((type-lib type-lib))
  (loop for index from 0 below (get-type-info-count type-lib)
     for type-info = (get-type-info type-lib index)
     for type-attr = (get-type-attr type-info)
     append (loop for i from 0 below (vars type-attr)
               collect (get-var-desc type-info i))))

(defmethod get-func-desc-list ((type-lib type-lib))
  (loop for index from 0 below (get-type-info-count type-lib)
     for type-info = (get-type-info type-lib index)
     for type-attr = (get-type-attr type-info)
     append (loop for i from 0 below (funcs type-attr)
               collect (get-func-desc type-info i))))
