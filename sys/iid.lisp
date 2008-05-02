;;;; -*- Mode: LISP; Syntax: COMMON-LISP; Coding: shift_jis; -*-
(in-package :cl-win32ole-sys)

(defun print-clsid (clsid)
  (format t "{~8,'0X-~4,'0x-~4,'0x-~2,'0x~2,'0x-"
          (cffi:foreign-slot-value clsid 'GUID 'Data1)
          (cffi:foreign-slot-value clsid 'GUID 'Data2)
          (cffi:foreign-slot-value clsid 'GUID 'Data3)
          (cffi:mem-aref (cffi:foreign-slot-value clsid 'GUID 'Data4)
                         :unsigned-char 0)
          (cffi:mem-aref (cffi:foreign-slot-value clsid 'GUID 'Data4)
                         :unsigned-char 1))
  (loop for i from 2 to 7
     do (format t "~2,'0x"
                (cffi:mem-aref (cffi:foreign-slot-value clsid 'GUID 'Data4)
                               :unsigned-char i)))
  (format t "}"))

(defun make-clsid (str)
  (let ((clsid (cffi:foreign-alloc 'GUID)))
    (labels ((f (x)
               (parse-integer x :radix 16))
             (s (n x)
               (setf (cffi:mem-aref
                      (cffi:foreign-slot-value clsid 'GUID 'Data4)
                      :unsigned-char n) x)))
      (cl-ppcre:do-register-groups
          ((#'f d1) (#'f d2) (#'f d3)
           (#'f c0) (#'f c1) (#'f c2) (#'f c3)
           (#'f c4) (#'f c5) (#'f c6) (#'f c7))
          ("{(........)-(....)-(....)-(..)(..)-(..)(..)(..)(..)(..)(..)}" str)
        (setf (cffi:foreign-slot-value clsid 'GUID 'Data1) d1
              (cffi:foreign-slot-value clsid 'GUID 'Data2) d2
              (cffi:foreign-slot-value clsid 'GUID 'Data3) d3)
        (s 0 c0)
        (s 1 c1)
        (s 2 c2)
        (s 3 c3)
        (s 4 c4)
        (s 5 c5)
        (s 6 c6)
        (s 7 c7)))
    clsid))


(defvar IID_IDispatch (make-clsid "{00020400-0000-0000-C000-000000000046}"))

(defvar IID_NULL (make-clsid "{00000000-0000-0000-0000-000000000000}"))
