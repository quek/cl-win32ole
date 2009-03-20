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
      (list :date
            (cffi:foreign-slot-value st 'SYSTEMTIME 'year)
            (cffi:foreign-slot-value st 'SYSTEMTIME 'month)
            (cffi:foreign-slot-value st 'SYSTEMTIME 'day)
            (cffi:foreign-slot-value st 'SYSTEMTIME 'hour)
            (cffi:foreign-slot-value st 'SYSTEMTIME 'minute)
            (cffi:foreign-slot-value st 'SYSTEMTIME 'second)
            (cffi:foreign-slot-value st 'SYSTEMTIME 'millisecond)))))
