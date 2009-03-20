(in-package :cl-win32ole)

(defclass bstr ()
  ((ptr :accessor ptr)
   (string :initarg :string)))

(defclass variant ()
  ((ptr :initarg :ptr :accessor ptr)))

(defclass unknown ()
  ((ptr :initarg :ptr :accessor ptr)
   (finalizer :initarg :finalizer :initform nil)))

(defclass dispatch (unknown)
  ())

(defclass type-info (unknown)
  ())

(defclass type-lib (unknown)
  ())

(defclass type-attr ()
  ((funcs :initarg :funcs :reader funcs)
   (vars :initarg :vars :reader vars)
   (imple-types :initarg :imple-types :reader imple-types)))

(defclass desc-mixin ()
  ((name :initarg :name :reader name)
   (memid :initarg :memid :reader memid)))

(defclass func-desc (desc-mixin)
  ((invkind :initarg :invkind :reader invkind)))

(defclass var-desc (desc-mixin)
  ((value :initarg :value :reader value)
   (varkind :initarg :varkind :reader varkind)
   (var-flags :initarg :var-flags :reader var-flags)))


(defgeneric get-type-info (object &optional index))
