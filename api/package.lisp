;;;; -*- Mode: LISP; Syntax: COMMON-LISP; -*-
(in-package :cl-user)

(defpackage #:cl-win32ole
  (:nicknames #:win32ole #:ole)
  (:use #:common-lisp #:cl-win32ole-sys
        #+sbcl #:sb-mop)
  (:export #:with-co-initialize
           #:create-object
           #:property
           #:invoke
           #:ole
           #:with-ole-object
	   #:empty-array))
