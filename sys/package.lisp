;;;; -*- Mode: LISP; Syntax: COMMON-LISP; -*-
(in-package :cl-user)

(defpackage #:cl-win32ole-sys
  (:nicknames #:win32ole-sys #:ole-sys)
  (:use #:common-lisp)
  (:export #:succeeded
           #:with-ole-str

           #:VARIANT
           #:VT_EMPTY
           #:VT_NULL
           #:VT_I2
           #:VT_I4
           #:VT_R4
           #:VT_R8
           #:VT_CY
           #:VT_DATE
           #:VT_BSTR
           #:VT_DISPATCH
           #:VT_ERROR
           #:VT_BOOL
           #:VT_VARIANT
           #:VT_UNKNOWN
           #:VT_DECIMAL
           #:VT_I1
           #:VT_UI1
           #:VT_UI2
           #:VT_UI4
           #:VT_I8
           #:VT_UI8
           #:VT_INT
           #:VT_UINT
           #:VT_VOID
           #:VT_HRESULT
           #:VT_PTR
           #:VT_SAFEARRAY
           #:VT_CARRAY
           #:VT_USERDEFINED
           #:VT_LPSTR
           #:VT_LPWSTR
           #:VT_RECORD
           #:VT_INT_PTR
           #:VT_UINT_PTR
           #:VT_FILETIME
           #:VT_BLOB
           #:VT_STREAM
           #:VT_STORAGE
           #:VT_STREAMED_OBJECT
           #:VT_STORED_OBJECT
           #:VT_BLOB_OBJECT
           #:VT_CF
           #:VT_CLSID
           #:VT_VERSIONED_STREAM
           #:VT_BSTR_BLOB
           #:VT_VECTOR
           #:VT_ARRAY
           #:VT_BYREF
           #:VT_RESERVED
           #:VT_ILLEGAL
           #:VT_ILLEGALMASKED
           #:VT_TYPEMASK
           #:VARIANT_TRUE
           #:VARIANT_FALSE
           #:alloc-variant
           #:free-variant
           #:lisp->variant
           #:bstr->lisp
           #:lisp->bstr
           #:variant-type
           #:variant-type*
           #:variant-value
           #:variant-array-p
           #:variant-byref-p
           #:variant-array-value
           #:VariantInit
           #:VariantClear

           #:create-instance

           #:unknown-add-ref
           #:unknown-release

           #:dispatch-add-ref
           #:dispatch-release
           #:dispatch-get-ids-of-names
           #:dispatch-get-property
           #:dispatch-put-property
           #:dispatch-invoke
           #:dispatch-get-type-info-count
           #:dispatch-get-type-info

           #:type-info-add-ref
           #:type-info-release
           #:type-info-get-type-attr
           #:type-info-get-func-desc
           #:type-info-get-var-desc
           #:type-info-get-ref-type-of-impl-type
           #:type-info-get-name-of-documentation
           #:type-info-get-containing-type-lib
           #:type-info-get-ref-type-info
           #:type-info-release-type-attr
           #:type-info-release-func-desc
           #:type-info-release-var-desc

           #:type-lib-get-type-info-count
           #:type-lib-get-name-of-documentation
           #:type-lib-get-type-info

           #:TYPEATTR
           #:cFuncs
           #:cVars
           #:cImplTypes

           #:FUNCDESC
           #:memid
           #:invkind

           #:VARDESC
           #:memid
           #:value
           #:varkind
           #:wVarFlags

           #:VAR_PERINSTANCE
           #:VAR_STATIC
           #:VAR_CONST
           #:VAR_DISPATCH

           #:VARFLAG_FHIDDEN
           #:VARFLAG_FRESTRICTED
           #:VARFLAG_FNONBROWSABLE

           #:->safe-array
           #:safe-array->variant-ptr-list

           #:SYSTEMTIME
           #:UDATE
           #:VarUdateFromDate
           #:VarDateFromUdate
           #:from-variant-date
           #:to-variant-date

	   #:switch
	   #:eswitch

           #:p
           #:dformat
           ))

