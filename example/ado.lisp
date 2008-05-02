(eval-when (:compile-toplevel :load-toplevel :execute)
  (require :cl-win32ole)
  (use-package :cl-win32ole))

(defvar *src-dir* (format nil "~a:~a"
                        (pathname-device *load-truename*)
                        (directory-namestring *load-truename*)))
(defun csv-example ()
  (let ((cn (create-object "ADODB.Connection"))
        (rs (create-object "ADODB.Recordset"))
        (cs (format nil "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=~a;Extended Properties=\"text;HDR=Yes;FMT=Delimited;\";" *src-dir*)))
    (property cn :ConnectionString cs)
    (invoke cn :open)
    (invoke rs :open "select * from aaa.csv" cn)
    (unwind-protect
         (loop until (property rs :eof)
            do (progn
                 (format t "key: ~a, value: ~a, etc: ~a~%"
                         (property (invoke (property rs :fields)
                                           :item "key") :value)
                         (property (invoke (property rs :fields)
                                           :item "value") :value)
                         (property (invoke (property rs :fields)
                                           :item "etc") :value))
                 (invoke rs :movenext))))
      (invoke rs :close)
      (invoke cn :close)))

(defvar *mdb-path* (merge-pathnames "test.mdb" *src-dir*))

(defvar *conn-str* (format nil "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=~a;Jet OLEDB:Engine Type=5;" *mdb-path*))

(defun create-test-mdb ()
  (let ((adox (create-object "ADOX.Catalog")))
    (invoke adox :create *conn-str*)))

(defun mdb-example ()
  (unless (probe-file *mdb-path*)
    (create-test-mdb))
  (let ((cn (create-object "ADODB.Connection")))
    (invoke cn :open *conn-str*)
    (unwind-protect
         (progn
           (invoke cn :execute
                   "create table table1 (
                      str1 number(10) primary key,
                      str varchar(10))")
           )
      (invoke cn :close))))

;;(csv-example)
;;(mdb-example)