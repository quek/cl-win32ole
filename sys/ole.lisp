;;;; -*- Mode: LISP; Syntax: COMMON-LISP; -*-
(in-package :cl-win32ole-sys)


(defvar *lcid* 2048)                    ;LOCALE_SYSTEM_DEFAULT

(defvar *co-initialized* nil)

(unless *co-initialized*
  (format t "co-initialize~%")
  (succeeded (co-initialize (cffi-sys:null-pointer)))
  (setf *co-initialized* t))


(defmacro with-co-initialize (&body body)
  `(progn
     (succeeded (co-initialize (cffi-sys:null-pointer)))
     (unwind-protect
          (progn ,@body)
       (co-uninitialize))))


(defun message-from-system (HRESULT)
  (cffi:with-foreign-object (buffer :unsigned-short 4096)
    (let ((ret (FormatMessageW
                (+ FORMAT_MESSAGE_FROM_SYSTEM
                   FORMAT_MESSAGE_IGNORE_INSERTS)
                (cffi-sys:null-pointer)
                HRESULT
                *lcid*
                buffer
                1024
                (cffi-sys:null-pointer))))
      (with-output-to-string (*standard-output*)
        (loop for i from 0 below ret
           do (let ((c (code-char
                        (cffi:mem-aref buffer :unsigned-short i))))
                (unless (member c '(#\Return #\Linefeed))
                  (write-char c))))))))


(defun create-instance (prog-id)
  (cffi:with-foreign-objects ((clsid 'CLSID))
    (with-ole-str (s prog-id)
      (succeeded (clsid-from-prog-id s clsid)))
    (cffi:with-foreign-object (pdispatch :pointer)
      (succeeded
       (co-create-instance clsid
                           (cffi-sys:null-pointer)
                           (+ CLSCTX_INPROC_SERVER CLSCTX_LOCAL_SERVER)
                           IID_IDispatch
                           pdispatch))
      (cffi:mem-aref pdispatch :pointer))))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; type-lib
(defun type-info-get-type-attr (type-info)
  (cffi:with-foreign-object (type-attr :pointer)
    (succeeded (%type-info-get-type-attr type-info type-attr))
    (cffi:mem-aref type-attr :pointer)))

(defun type-info-get-func-desc (type-info index)
  (cffi:with-foreign-object (func-desc :pointer)
    (succeeded (%type-info-get-func-desc type-info index func-desc))
    (cffi:mem-aref func-desc :pointer)))

(defun type-info-get-var-desc (type-info index)
  (cffi:with-foreign-object (var-desc :pointer)
    (succeeded (%type-info-get-var-desc type-info index var-desc))
    (cffi:mem-aref var-desc :pointer)))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; type-info
(defun type-info-get-ref-type-of-impl-type (type-info index)
  (cffi:with-foreign-object (ref-type 'DWORD)
    (succeeded (%type-info-get-ref-type-of-impl-type type-info index ref-type))
    (cffi:mem-aref ref-type 'DWORD)))

(defun type-info-get-name-of-documentation (type-info &optional (index -1))
  (cffi:with-foreign-object (name :pointer)
    (succeeded (%type-info-get-documentation type-info
                                    index
                                    name
                                    (cffi-sys:null-pointer)
                                    (cffi-sys:null-pointer)
                                    (cffi-sys:null-pointer)))
    (let ((bstr (cffi:mem-aref name :pointer)))
      (prog1 (bstr->lisp bstr)
             (SysFreeString bstr)))))

(defun type-info-get-containing-type-lib (type-info)
  (cffi:with-foreign-objects ((type-lib :pointer)
                              (index :pointer))
    (succeeded (%type-info-get-containing-type-lib type-info
                                                   type-lib
                                                   index))
    (values (cffi:mem-aref type-lib :pointer)
            (cffi:mem-aref index 'UINT))))

(defun type-info-get-ref-type-info (type-info ref-type)
  (cffi:with-foreign-object (ref-type-info :pointer)
    (succeeded (%type-info-get-ref-type-info type-info ref-type ref-type-info))
    (cffi:mem-aref ref-type-info :pointer)))

(defun type-lib-get-name-of-documentation (type-lib index)
  (cffi:with-foreign-object (name :pointer)
    (succeeded (%type-lib-get-documentation type-lib
                                            index
                                            name
                                            (cffi-sys:null-pointer)
                                            (cffi-sys:null-pointer)
                                            (cffi-sys:null-pointer)))
    (let ((bstr (cffi:mem-aref name :pointer)))
      (prog1 (bstr->lisp bstr)
        (SysFreeString bstr)))))

(defun type-lib-get-type-info (dispatch index)
  (cffi:with-foreign-object (pp-type-info :pointer)
    (succeeded (%type-lib-get-type-info dispatch
                                        index
                                        pp-type-info))
    (cffi:mem-aref pp-type-info :pointer)))
