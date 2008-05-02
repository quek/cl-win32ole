(in-package :cl-win32ole)

(defun make-dispatch (ptr)
  (make-instance 'dispatch :ptr ptr))

(defmethod add-ref ((dispatch dispatch))
  (unknown-add-ref (ptr dispatch)))

(defmethod release ((dispatch dispatch))
  (unknown-release (ptr dispatch)))

(defmethod get-ids-of-names ((dispatch dispatch) name)
  (dispatch-get-ids-of-names dispatch (string name)))

(defmethod get-property ((dispatch dispatch) property)
  (to-lisp (make-instance 'variant
                          :ptr (dispatch-get-property (ptr dispatch)
                                                      (string property)))))

(defmethod put-property ((dispatch dispatch) property value)
  (dispatch-put-property
   (ptr dispatch) (string property) (ptr (make-variant value)))
  value)

(defmethod property ((dispatch dispatch) property
                     &optional (value nil value-p))
  (if value-p
      (put-property dispatch property value)
      (get-property dispatch property)))

(defmethod invoke ((dispatch dispatch) method &rest args)
  (let ((variant-ptr (apply #'dispatch-invoke (ptr dispatch) (string method)
                            (mapcar #'(lambda (x)
                                        (ptr (make-variant x)))
                                    args))))
    (to-lisp (make-instance 'variant :ptr variant-ptr))))

(defmethod get-type-info-count ((dispatch dispatch))
  (dispatch-get-type-info-count (ptr dispatch)))

(defmethod %get-type-info ((dispatch dispatch))
  (make-type-info (dispatch-get-type-info (ptr dispatch))))

(defmethod get-type-info ((dispatch dispatch) &optional index)
  (declare (ignore index))
  (let* ((type-info (%get-type-info dispatch))
         (name (get-name-of-documentation type-info))
         (type-lib (get-containing-type-lib type-info))
         (type-info-count (get-type-info-count type-lib)))
    (loop for i from 0 below type-info-count
       for nm = (get-name-of-documentation type-lib i)
       when (string= name nm)
       do (return (get-type-info type-lib i)))))

;; TODO can't work!!
(defmethod methods ((dispatch dispatch))
  (let ((type-lib (get-containing-type-lib (get-type-info dispatch))))
    (mapcar #'name (get-func-desc-list type-lib))))
;;  (let* ((type-info (get-type-info dispatch))
;;         (ref-type-info-list (get-ref-type-info-list type-info)))
;;    (loop for each in ref-type-info-list
;;       append (loop for i from 0 below (funcs (get-type-attr each))
;;                 collect (let ((func-desc (get-func-desc each i)))
;;                           (get-name-of-documentation
;;                            each (memid func-desc)))))))

(defmethod get-var-desc-list ((dispatch dispatch))
  (get-var-desc-list (get-containing-type-lib (%get-type-info dispatch))))

(defmethod load-constants ((dispatch dispatch)
                           &key (package *package*)
                           make-package-p)
  (let ((constants (loop for var-desc in (get-var-desc-list dispatch)
                      when (loadable-const-p var-desc)
                      collect (cons (name var-desc)
                                   (value var-desc)))))
    (when make-package-p
      (make-package package))
    (def-ole-constants constants package)))

(defun def-ole-constants (constants package)
  (loop for (name . var) in constants
     with p = (find-package package)
     for sym = (intern name p)
     do (progn (eval (list 'defconstant sym var))
               (export sym p))))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; property access
(defmethod slot-missing ((class (eql (find-class 'dispatch)))
                         instance slot-name (operation (eql 'slot-value))
                         &optional new-value)
  (declare (ignore class new-value))
  (property instance slot-name))

(defmethod slot-missing ((class (eql (find-class 'dispatch)))
                         instance slot-name (operation (eql 'setf))
                         &optional new-value)
  (declare (ignore class))
  (property instance slot-name new-value))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; create object
(defmethod create-object ((prog-id string))
  (make-instance 'dispatch :ptr (create-instance prog-id)))

(defmacro with-ole-object ((var prog-id) &body body)
  (let ((original-function (gensym)))
    `(let ((,var (create-object ,prog-id))
           (,original-function (and (fboundp ',var) (symbol-function ',var))))
       (setf (symbol-function ',var) ,var)
       (unwind-protect
            (progn
              ,@body)
         (if ,original-function
             (setf (symbol-function ',var) ,original-function)
             (fmakunbound ',var))))))