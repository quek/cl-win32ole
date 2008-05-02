;;;; -*- Mode: LISP; Syntax: COMMON-LISP; -*-
(in-package :cl-win32ole-sys)

(cffi:defcfun "FormatMessageW" DWORD
  (dwFlangs DWORD)
  (lpSource :pointer)
  (dwMessageId DWORD)
  (dwLanguageId DWORD)
  (lpBuffer :pointer)
  (nSize DWORD)
  (Arguments :pointer))

(cffi:defcfun ("CoInitialize" co-initialize) HRESULT
  (arg :pointer))

(cffi:defcfun ("CoUninitialize" co-uninitialize) :void)


(cffi:defcfun ("CLSIDFromProgID" clsid-from-prog-id) HRESULT
  (spszProgID :string)
  (pclsid :pointer))

(cffi:defcfun ("CoCreateInstance" co-create-instance) HRESULT
  (rclsid :pointer)
  (pUnkOuter :pointer)
  (dw-cls-context :int16)
  (riid :pointer)
  (ppv :pointer))


(defun %unknown-function (unknown symbol)
  (cffi:foreign-slot-value
   (cffi:foreign-slot-value unknown 'IUnknown 'vtbl)
   'IUnknownVtbl symbol))

(defun unknown-add-ref (unknown)
  (cffi:foreign-funcall-pointer
   (%unknown-function unknown 'AddRef) ()
   :pointer unknown
   :unsigned-long))

(defun unknown-release (unknown)
  (cffi:foreign-funcall-pointer
   (%unknown-function unknown 'Release) ()
   :pointer unknown
   :unsigned-long))


(defun %dispatch-function (dispatch symbol)
  (cffi:foreign-slot-value
   (cffi:foreign-slot-value dispatch 'IDispatch 'vtbl)
   'IDispatchVtbl symbol))

(defun %dispatch-get-ids-of-names (dispatch riid rgszNames cNames lcid rgDispId)
  (cffi:foreign-funcall-pointer
   (%dispatch-function dispatch 'GetIDsOfNames) ()
   :pointer dispatch
   :pointer riid
   :pointer rgszNames
   UINT cNames
   LCID lcid
   :pointer rgDispId
   HRESULT))

(defun %dispatch-invoke (dispatch
                         dispIdMember
                         riid
                         lcid
                         wFlags
                         pDispParams
                         pVarResult
                         pExcepInfo
                         puArgErr)
  (cffi:foreign-funcall-pointer
   (%dispatch-function dispatch 'Invoke) ()
   :pointer dispatch
   DISPID dispIdMember
   :pointer riid
   LCID lcid
   WORD wFlags
   :pointer pDispParams
   :pointer pVarResult
   :pointer pExcepInfo
   :pointer puArgErr
   HRESULT))


(defun %dispatch-get-type-info-count (dispatch petinfo)
  (cffi:foreign-funcall-pointer
   (%dispatch-function dispatch 'GetTypeInfoCount) ()
   :pointer dispatch
   :pointer petinfo
   HRESULT))

(defun %dispatch-get-type-info (dispatch iTInfo lcid ppTInof)
  (cffi:foreign-funcall-pointer
   (%dispatch-function dispatch 'GetTypeInfo) ()
   :pointer dispatch
   UINT iTInfo
   LCID lcid
   :pointer ppTInof
   HRESULT))


(defun %type-info-function (type-info symbol)
  (cffi:foreign-slot-value
   (cffi:foreign-slot-value type-info 'ITypeInfo 'vtbl)
   'ITypeInfoVtbl symbol))

(defun %type-info-get-type-attr (type-info ppTypeAttr)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'GetTypeAttr) ()
   :pointer type-info
   :pointer ppTypeAttr
   HRESULT))

(defun %type-info-get-func-desc (type-info index ppFuncDesc)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'GetFuncDesc) ()
   :pointer type-info
   UINT index
   :pointer ppFuncDesc
   HRESULT))

(defun %type-info-get-var-desc (type-info index ppVarDesc)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'GetVarDesc) ()
   :pointer type-info
   UINT index
   :pointer ppVarDesc
   HRESULT))

(defun %type-info-get-ref-type-of-impl-type (type-info index pRefType)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'GetRefTypeOfImplType) ()
   :pointer type-info
   UINT index
   :pointer pRefType
   HRESULT))

(defun %type-info-get-documentation (type-info memid pBstrName pBstrDocString
                                     pdwHelpContext pBstrHelpFile)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'GetDocumentation) ()
   :pointer type-info
   MEMBERID memid
   :pointer pBstrName
   :pointer pBstrDocString
   :pointer pdwHelpContext
   :pointer pBstrHelpFile
   HRESULT))

(defun %type-info-get-containing-type-lib (type-info ppTLib pIndex)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'GetContainingTypeLib) ()
   :pointer type-info
   :pointer ppTLib
   :pointer pIndex
   HRESULT))

(defun %type-info-get-ref-type-info (type-info ref-type ppTInfo)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'GetRefTypeInfo) ()
   :pointer type-info
   HREFTYPE ref-type
   :pointer ppTInfo
   HRESULT))

(defun type-info-release-type-attr (type-info type-attr)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'ReleaseTypeAttr) ()
   :pointer type-info
   :pointer type-attr
   :void))

(defun type-info-release-func-desc (type-info func-desc)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'ReleaseFuncDesc) ()
   :pointer type-info
   :pointer func-desc
   :void))

(defun type-info-release-var-desc (type-info var-desc)
  (cffi:foreign-funcall-pointer
   (%type-info-function type-info 'ReleaseVarDesc) ()
   :pointer type-info
   :pointer var-desc
   :void))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; type-lib
(defun %type-lib-function (type-lib symbol)
  (cffi:foreign-slot-value
   (cffi:foreign-slot-value type-lib 'ITypeLib 'vtbl)
   'ITypeLibVtbl symbol))

(defun type-lib-add-ref (type-lib)
  (cffi:foreign-funcall-pointer
   (%type-lib-function type-lib 'AddRef) ()
   :pointer type-lib
   :unsigned-long))

(defun type-lib-release (type-lib)
  (cffi:foreign-funcall-pointer
   (%type-lib-function type-lib 'Release) ()
   :pointer type-lib
   :unsigned-long))

(defun type-lib-get-type-info-count (type-lib)
  (cffi:foreign-funcall-pointer
   (%type-lib-function type-lib 'GetTypeInfoCount) ()
   :pointer type-lib
   UINT))

(defun %type-lib-get-documentation (type-lib index pBstrName pBstrDocString
                                    pdwHelpContext pBstrHelpFile)
  (cffi:foreign-funcall-pointer
   (%type-lib-function type-lib 'GetDocumentation) ()
   :pointer type-lib
   :int index
   :pointer pBstrName
   :pointer pBstrDocString
   :pointer pdwHelpContext
   :pointer pBstrHelpFile
   HRESULT))

(defun %type-lib-get-type-info (type-lib index ppTInof)
  (cffi:foreign-funcall-pointer
   (%type-lib-function type-lib 'GetTypeInfo) ()
   :pointer type-lib
   UINT index
   :pointer ppTInof
   HRESULT))


(cffi:defcfun "SysAllocString" :pointer ;BSTR
  (ole-str :pointer))                   ;const OLECHAR *

(cffi:defcfun "SysFreeString" :void
  (bstr :pointer))                      ;BSTR

(cffi:defcfun "VariantInit" :void
  (pvarg :pointer))                     ;VARIANTARG * pvarg

(cffi:defcfun "VariantClear" HRESULT
  (pvarg :pointer))                     ;VARIANTARG * pvarg

(cffi:defcfun "VariantCopy" HRESULT
  (pvargDest :pointer)                  ;VARIANTARG * pvargDest
  (pvargSrc :pointer))                  ;VARIANTARG * pvargSrc


(cffi:defcfun "SafeArrayGetDim" UINT
  (psa :pointer))

(cffi:defcfun "SafeArrayGetUBound" HRESULT
  (pas :pointer)
  (dim UINT)
  (u-bound :pointer))                   ; LONG *

(cffi:defcfun "SafeArrayGetLBound" HRESULT
  (pas :pointer)
  (dim UINT)
  (l-bound :pointer))                   ; LONG *

(cffi:defcfun "SafeArrayPtrOfIndex" HRESULT
  (pas :pointer)
  (indices :pointer)                    ; LONG *
  (data :pointer))                      ; void **

(cffi:defcfun "SafeArrayLock" HRESULT
  (psa :pointer))

(cffi:defcfun "SafeArrayUnlock" HRESULT
  (psa :pointer))

(cffi:defcfun "SafeArrayCreate" :pointer ; SAFEARRAY *
  (vt VARTYPE)
  (dims UINT)
  (sabound :pointer))                   ; SAFEARRAYBOUND *

(cffi:defcfun "SafeArrayPutElement" HRESULT
  (psa :pointer)                        ; SAFEARRAY *
  (indices :pointer)                    ; LONG *
  (v :pointer))                         ; void *

(cffi:defcfun "SafeArrayGetElement" HRESULT
  (psa :pointer)                        ; SAFEARRAY *
  (indices :pointer)                    ; LONG *
  (v :pointer))                         ; void *




(defun HRESULT->DWORD (HRESULT)
  (* -1 (1+ (logxor #xffffffff HRESULT))))

(defmacro succeeded (form)
  `(let ((hresult ,form))
     (when (< hresult #x00000000)
       (let ((dwMessageId (HRESULT->DWORD hresult)))
         (error "ERROR ~A(~X)~%~A~%"
                (message-from-system dwMessageId) dwMessageId ',form)))))
