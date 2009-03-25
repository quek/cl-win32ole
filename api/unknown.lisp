(in-package :cl-win32ole)

(defmethod initialize-instance :after ((unknown unknown) &rest args)
  (declare (ignore args))
  (let ((ptr (ptr unknown))
        (finalizer (slot-value unknown 'finalizer))
        (class (class-name (class-of unknown))))
    (dformat t "unknown::alloc ~a~%" ptr)
    (finalize unknown (lambda ()
                        (dformat t "~a::realease ~a~%" class ptr)
                        (when finalizer
                          (free-variant
                           (apply #'dispatch-invoke ptr
                                  (if (atom finalizer)
                                      (list (string finalizer))
                                      (cons (string (car finalizer))
                                            (mapcar #'(lambda (x)
                                                        (ptr (make-variant x)))
                                                    (cdr finalizer)))))))
                        (unknown-release ptr)))))

(defmethod add-ref ((unknown unknown))
  (unknown-add-ref (ptr unknown)))

(defmethod release ((unknown unknown))
  (unknown-release (ptr unknown)))

