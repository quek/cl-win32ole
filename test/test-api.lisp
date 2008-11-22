(eval-when (:compile-toplevel :load-toplevel :execute)
  (asdf:oos 'asdf:load-op :cl-win32ole)
  (asdf:oos 'asdf:load-op :fiveam))

(in-package :cl-win32ole)

(setf 5am:*debug-on-error* t
      5am:*debug-on-failure* t)

(5am:def-suite cl-win32ole-test)

(5am:in-suite cl-win32ole-test)

(5am:test variant-test
  (5am:is (string= "Hello" (to-lisp (make-variant "Hello"))))
  (5am:is (eq t (to-lisp (make-variant t))))
  (5am:is (eq nil (to-lisp (make-variant nil))))
  (5am:is (= 123 (to-lisp (make-variant 123))))
  (5am:is (= -123 (to-lisp (make-variant -123))))
  )

(5am:test ie-test
  (let* ((ie (create-object "InternetExplorer.Application"))
         (type-info (%get-type-info ie)))
    (unwind-protect
         (progn
           (5am:is-false (property ie :visible))
           (property ie :visible t)
           (5am:is-true (property ie :visible))
           (invoke ie :navigate "http://www.cliki.net/index")
           (loop while (property ie :busy) do (sleep 0.5))
           (5am:is (string= "CLiki : index"
                            (property (property ie :document) :title)))

           (multiple-value-bind (type-lib index)
               (get-containing-type-lib type-info)
             (5am:is (= 4 index))
             (5am:is (= 30 (get-type-info-count type-lib))))
           )
      (invoke ie :quit))))

(5am:test wsh-test
  (let* ((wsh (create-object "WScript.Shell"))
         (type-info (%get-type-info wsh)))
    (5am:is (string= "IWshShell3" (get-name-of-documentation type-info)))
    (multiple-value-bind (type-lib index) (get-containing-type-lib type-info)
      (5am:is (= 5 index))
      (5am:is (= 55 (get-type-info-count type-lib)))
      )
    (invoke wsh :popup "Hello Lisp!" 1 "cl-win32ole" 32)
    ))

(5am:test type-lib-test
  (let* ((wsh (create-object "WScript.Shell"))
         (type-info (get-type-info wsh)))
    (5am:is (string= "IWshShell3" (get-name-of-documentation type-info)))))

(5am:test test-slot-missing
  (let ((ie (create-object "InternetExplorer.Application")))
    (unwind-protect
         (with-slots (visible busy document) ie
           (setf visible t)
           (invoke ie :navigate "http://www.cliki.net/index")
           (loop while busy do (sleep 0.5))
           (5am:is (string= "CLiki : index" (property document :title))))
      (invoke ie 'quit))))

(5am:test test-methods
  (let ((ie (create-object "InternetExplorer.Application")))
    ;; TODO now not work well!
    (p (methods ie))))


(5am:test test-load-constants
  (labels ((? (s v)
             (5am:is (= (symbol-value (intern s (find-package :ie-const)))
                        v))))
    (make-package :ie-const)
    (unwind-protect
         (let ((ie (create-object "InternetExplorer.Application")))
           (load-constants ie :package :ie-const)
           (? "READYSTATE_UNINITIALIZED" 0)
           (? "READYSTATE_LOADING" 1)
           (? "READYSTATE_LOADED" 2)
           (? "READYSTATE_INTERACTIVE" 3)
           (? "READYSTATE_COMPLETE" 4))
      (delete-package :ie-const))))

(5am:test excel-test
  (let ((ex (create-object "Excel.Application")))
    (with-slots (visible workbooks) ex
      (setf visible t)
      (let* ((book (ole workbooks :add ))
             (sheet (ole book :worksheets :item 1)))
        (setf (ole sheet :range "A1:C2" :value)
              '(("Noth" "South" "Quis") (5.2 10 300)))
        (5am:is (equalp '(("Noth" "South" "Quis") (5.2 10 300))
                        (ole sheet :range "A1:C2" :value)))
        (setf (ole book 'saved) t)))
    (invoke ex :quit)))

(5am:run! 'cl-win32ole-test)
