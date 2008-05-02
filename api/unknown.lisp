(in-package :cl-win32ole)

(defmethod initialize-instance :after ((unknown unknown) &rest args)
  (declare (ignore args))
  (let ((ptr (ptr unknown))
        (class (class-name (class-of unknown))))
    (format t "unknown::alloc ~a~%" ptr)
    (finalize unknown (lambda ()
                        (format t "~a::realease ~a~%" class ptr)
                        (unknown-release ptr)))))

(defmethod add-ref ((unknown unknown))
  (unknown-add-ref (ptr unknown)))

(defmethod release ((unknown unknown))
  (unknown-release (ptr unknown)))

