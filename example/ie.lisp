(eval-when (:compile-toplevel :load-toplevel :execute)
  (require :cl-win32ole)
  (use-package :cl-win32ole))

(defun ie-example1 ()
  (let ((ie (create-object "InternetExplorer.Application")))
    (with-slots (visible busy document) ie
      (setf visible t)
      (funcall ie :navigate "http://sbcl.sourceforge.net/")
      (loop while busy
         do (sleep 0.5))
      (format t "document title is \"~a\".~%" (slot-value document :title))
      (sleep 3))
    (funcall ie :quit)))

(defun ie-example2 ()
  (let ((ie (create-object "InternetExplorer.Application")))
    (setf (ole ie :visible) t)
    (ole ie :navigate "http://www.google.co.jp/")
    (loop while (ole ie :busy) do (sleep 0.5))
    (setf (ole ie :document :all :item "q" :value) "Common Lisp")
    (ole ie :document :all :item "btnG" :click)
    (sleep 3)
    (ole ie :quit)))

(defun ie-example3 ()
  (let ((ie (create-object "InternetExplorer.Application")))
    (with-slots (visible busy document) ie
      (setf visible t)
      (invoke ie :navigate "http://www.yahoo.co.jp/")
      (loop while busy do (sleep 0.5))
      (with-slots (all) document
        (p (invoke all :tags "A")))))
  )
