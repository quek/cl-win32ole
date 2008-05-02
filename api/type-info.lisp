(in-package :cl-win32ole)

(defun make-type-info (ptr)
  (make-instance 'type-info :ptr ptr))

(defmethod get-name-of-documentation ((type-info type-info)
                                      &optional (index -1))
  (type-info-get-name-of-documentation (ptr type-info) index))

(defmethod get-containing-type-lib ((type-info type-info))
  (multiple-value-bind (type-lib index)
      (type-info-get-containing-type-lib (ptr type-info))
    (values (make-type-lib type-lib) index)))

(defmethod get-type-attr ((type-info type-info))
  (let ((ptr-type-attr (type-info-get-type-attr (ptr type-info))))
    (unwind-protect
         (make-type-attr ptr-type-attr)
      (type-info-release-type-attr (ptr type-info) ptr-type-attr))))

(defmethod get-ref-type-info-list ((type-info type-info))
  (let ((type-attr (get-type-attr type-info)))
    (loop for i from 0 below (imple-types type-attr)
       collect (make-type-info
                (let ((ref-type (type-info-get-ref-type-of-impl-type
                                 (ptr type-info) i)))
                  (type-info-get-ref-type-info (ptr type-info) ref-type))))))

(defmethod get-func-desc ((type-info type-info) index)
  (let ((ptr-func-desc (type-info-get-func-desc (ptr type-info) index)))
    (unwind-protect
         (make-func-desc ptr-func-desc type-info)
      (type-info-release-func-desc (ptr type-info) ptr-func-desc))))

(defmethod get-var-desc ((type-info type-info) index)
  (let ((ptr-var-desc (type-info-get-var-desc (ptr type-info) index)))
    (unwind-protect
         (make-var-desc ptr-var-desc type-info)
      (type-info-release-var-desc (ptr type-info) ptr-var-desc))))
