(in-package :cl-win32ole)

(defmethod initialize-instance :after ((variant variant) &rest args)
  (declare (ignore args))
  (unless (slot-boundp variant 'ptr)
    (setf (slot-value variant 'ptr) (alloc-variant)))
  (let ((ptr (ptr variant)))
    (finalize variant (lambda ()
                        (free-variant ptr)))))

(defmethod to-lisp (ptr)
  (let ((v (variant-value ptr))
        (type (variant-type ptr)))
    (cond ((variant-array-p ptr)
           (variant-array-to-lisp ptr))
          (t (eswitch (logand type (lognot VT_BYREF))
               (VT_EMPTY nil)
               (VT_I4 v)
               (VT_R4 v)
               (VT_R8 v)
               (VT_DATE (to-lisp-date ptr))
               (VT_BSTR (bstr->lisp v))
               (VT_BOOL (= VARIANT_TRUE v))
               (VT_DISPATCH (let ((dispatch (make-dispatch v)))
			      (add-ref dispatch)
			      dispatch)))))))

(defun to-lisp-date (ptr)
  (from-variant-date ptr))

(defun map-dim (fn list)
  (if (consp (car list))
      (cons (map-dim fn (car list))
            (map-dim fn (cdr list)))
      (mapcar fn list)))

(defun variant-array-to-lisp (ptr)
  (let ((ptr-list (safe-array->variant-ptr-list ptr)))
    (map-dim #'(lambda (x)
                 (prog1 (to-lisp x)
                   (free-variant x)))
             ptr-list)))

(defmethod to-lisp ((variant variant))
  (to-lisp (ptr variant)))

(defmethod make-variant (lisp-value &optional variant-type)
  (make-instance 'variant :ptr (lisp->variant lisp-value variant-type)))

(defmethod make-variant ((lisp-value dispatch) &optional variant-type)
  (declare (ignore variant-type))
  (make-instance 'variant :ptr (lisp->variant (ptr lisp-value) VT_DISPATCH)))

(defmethod make-variant ((lisp-value (eql nil)) &optional variant-type)
  (make-instance 'variant :ptr (lisp->variant lisp-value variant-type)))

(defmethod make-variant ((lisp-value list) &optional variant-type)
  (declare (ignore variant-type))
  (let* ((variant-list (map-dim #'make-variant lisp-value))
         (ptr-list (map-dim #'ptr variant-list))
         (psa (->safe-array ptr-list)))
    (make-variant psa (logior VT_VARIANT VT_ARRAY))))

(defmethod make-variant ((lisp-value (eql 'empty-array)) &optional variant-type)
  (declare (ignore variant-type))
  (make-variant (->safe-array nil) (logior VT_VARIANT VT_ARRAY)))

(defun empty-array ()
  'empty-array)

(defmethod make-variant ((lisp-value dt:date-time) &optional variant-type)
  (declare (ignore variant-type))
  (make-instance 'variant :ptr (to-variant-date lisp-value)))


(defmethod print-object ((variant variant) stream)
  (print-unreadable-object (variant stream :type t :identity t)
    (dformat stream "vt:~x ~a" (variant-type (ptr variant)) (to-lisp variant))))

