(in-package :cl-win32ole)

(defun split-ole-args (args &optional acc)
  (if (endp args)
      (progn
        (when (consp (car acc))
          (setf (car acc) (reverse (car acc))))
        (reverse acc))
      (progn
        (if (keywordp (car args))
            (progn
              (when (consp (car acc))
                (setf (car acc) (reverse (car acc))))
              (push (list (car args)) acc))
            (push (car args) (car acc)))
        (split-ole-args (cdr args) acc))))

(defun ole (dispatch &rest args)
  (let ((ole-args (split-ole-args args))
        (result dispatch))
    (dolist (i ole-args result)
      (setf result (apply #'invoke result i)))))

(defun (setf ole) (new-value dispatch &rest args)
  (let* ((butlast (butlast args))
         (last (last args))
         (object (apply #'ole dispatch butlast)))
    (apply #'property object (car last) (list new-value))))
