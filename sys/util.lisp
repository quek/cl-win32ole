(in-package :cl-win32ole-sys)

;; for debug
(defmacro p (&body body)
  `(progn ,@(mapcar #'(lambda (arg)
                        `(format t "~30a ; => ~a~%" ',arg ,arg))
                    body)))

(defmacro dformat (stream format &rest args)
  (declare (ignorable stream format args))
  ;;`(format ,stream ,format ,@args))
  )

(defmacro switch (keyform &body cases)
  "Switch is like case, except that it does not quote keys, and only accepts
one key per case."
  (let ((k (gensym)))
    `(let ((,k ,keyform))
       ,(reduce (lambda (case rest)
		  (destructuring-bind (key &body forms) case
		    `(cond ((eql ,key ,k) ,@forms)
			   (t ,rest))))
		cases
		:from-end t
		:initial-value nil))))
 
(defmacro eswitch (keyform &body cases)
  "ESwitch is like ecase, except that it does not quote keys, and only accepts
one key per case."
  (let ((k (gensym))
        (cases (mapcar (lambda (case)
                         (destructuring-bind (key &body forms) case
                           (cons (list (gensym) key) forms)))
                       cases)))
    `(let ((,k ,keyform)
           ,@(mapcar #'car cases))
       ,(reduce (lambda (case rest)
		  (destructuring-bind ((key-val key) &body forms) case
                    (declare (ignore key))
		    `(cond ((eql ,key-val ,k) ,@forms)
			   (t ,rest))))
		cases
		:from-end t
		:initial-value `(error "~a fell through ESWITCH expression. Wanted one of ~a"
				       ,k (list ,@(mapcar #'caar cases)))))))
	 
