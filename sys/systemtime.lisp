;;;; -*- Mode: LISP; Syntax: COMMON-LISP; -*-
(in-package :cl-win32ole-sys)

(cffi:defcstruct SYSTEMTIME
    (year WORD)
  (month WORD)
  (day-of-week WORD)
  (day WORD)
  (hour WORD)
  (minute WORD)
  (second WORD)
  (millisecond WORD))

(cffi:defcstruct UDATE
    (st SYSTEMTIME)
  (day-of-week USHORT))

(cffi:defcfun "VarUdateFromDate" HRESULT
  (date DATE)
  (flags :unsigned-long)
  (pudate :pointer))                    ; UDATE*

(cffi:defcfun "VarDateFromUdate" HRESULT
  (pudate :pointer)                     ; UDATE*
  (flags :unsigned-long)
  (pdate :pointer))                     ; DATE*(long*)


(defun from-variant-date (variant)
  (cffi:with-foreign-object (udate 'UDATE)
    (let ((date (if (variant-byref-p variant)
                    (cffi:mem-aref (variant-value variant) 'DATE)
                    (variant-value variant))))
      (succeeded (VarUdateFromDate date 0 udate)))
    (let ((st (cffi:foreign-slot-value udate 'UDATE 'st)))
      (make-instance
       'dt:date-time
       :year (cffi:foreign-slot-value st 'SYSTEMTIME 'year)
       :month (cffi:foreign-slot-value st 'SYSTEMTIME 'month)
       :day (cffi:foreign-slot-value st 'SYSTEMTIME 'day)
       :hour (cffi:foreign-slot-value st 'SYSTEMTIME 'hour)
       :minute (cffi:foreign-slot-value st 'SYSTEMTIME 'minute)
       :second (cffi:foreign-slot-value st 'SYSTEMTIME 'second)
       :millisecond (cffi:foreign-slot-value st 'SYSTEMTIME 'millisecond)))))

(defun to-variant-date (date-time)
  (cffi:with-foreign-objects ((pudate 'UDATE)
                              (pdate 'DATE))
    (let ((st (cffi:foreign-slot-value pudate 'UDATE 'st)))
      (macrolet
          ((m ()
             `(progn
                ,@(mapcar
                   (lambda (x)
                     `(setf (cffi:foreign-slot-value st 'SYSTEMTIME ',x)
                            (,(intern (concatenate 'string (symbol-name x)
                                                   "-OF") :dt)
                              date-time)))
                   '(year month day hour minute second millisecond)))))
        (m))
      (succeeded (VarDateFromUdate pudate 0 pdate))
      (let ((result (alloc-variant)))
        (VariantInit result)
        (setf (variant-type result) VT_DATE
              (variant-value result) (cffi:mem-aref pdate 'DATE))
        result))))
