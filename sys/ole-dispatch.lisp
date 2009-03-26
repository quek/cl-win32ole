(in-package :cl-win32ole-sys)

(defun dispatch-get-ids-of-names (dispatch name)
  (with-ole-str (ole-str name)
    (cffi:with-foreign-objects ((rgDispId :pointer)
                                (rgszNames :pointer))
      (setf (cffi:mem-aref rgszNames :pointer) ole-str)
      (succeeded (%dispatch-get-ids-of-names
                  dispatch
                  IID_NULL
                  rgszNames
                  1
                  *lcid*
                  rgDispId))
      (cffi:mem-ref rgDispId 'LONG))))

(defun dispatch-get-property (dispatch property)
  (let ((variant (alloc-variant)))
    (let ((disp-id (dispatch-get-ids-of-names dispatch property)))
      (%dispatch-invoke dispatch
                       disp-id
                       IID_NULL
                       *lcid*
                       DISPATCH_PROPERTYGET
                       *null-disp-params*
                       variant
                       (cffi-sys:null-pointer)
                       (cffi-sys:null-pointer)))
    variant))

(defun dispatch-put-property (dispatch property variant)
  (let ((disp-id (dispatch-get-ids-of-names dispatch property)))
    (cffi:with-foreign-objects
        ((args 'VARIANT)
         (params 'DISPPARAMS)
         (disp-id-named-args 'DISPID))
      (setf (cffi:mem-aref disp-id-named-args 'DISPID) DISPID_PROPERTYPUT)
      (VariantInit args)
      (variant-copy args variant)
      (setf (cffi:foreign-slot-value params 'DISPPARAMS 'rgvarg)
            args
            (cffi:foreign-slot-value params 'DISPPARAMS 'rgdispidNamedArgs)
            disp-id-named-args
            (cffi:foreign-slot-value params 'DISPPARAMS 'cArgs)
            1
            (cffi:foreign-slot-value params 'DISPPARAMS 'cNamedArgs)
            1)
      (succeeded (%dispatch-invoke dispatch
                                  disp-id
                                  IID_NULL
                                  *lcid*
                                  DISPATCH_PROPERTYPUT
                                  params
                                  (cffi-sys:null-pointer)
                                  (cffi-sys:null-pointer)
                                  (cffi-sys:null-pointer)))
      (VariantClear args)
      variant)))

;;(defun p-v (v)
;;  (format t "p-v: ~d ~x~%"
;;          (variant-type v)
;;          (variant-value v)))

(defun dispatch-invoke (dispatch method &rest args)
  (let ((disp-id (dispatch-get-ids-of-names dispatch method))
        (argc (length args))
        (reversed-args (reverse args)))
    (cffi:with-foreign-objects
        ((params 'DISPPARAMS)
         (v-args 'VARIANT argc)
         (excep-info 'EXCEPINFO)
         (arg-err :unsigned-int))
      (loop for i from 0 below argc
         for v = (cffi:mem-aref v-args 'VARIANT i)
         do (progn
              (VariantInit v)
              (variant-copy v (nth i reversed-args))))
      (setf (cffi:foreign-slot-value params 'DISPPARAMS 'rgvarg)
            v-args
            (cffi:foreign-slot-value params 'DISPPARAMS 'rgdispidNamedArgs)
            (cffi-sys:null-pointer)
            (cffi:foreign-slot-value params 'DISPPARAMS 'cArgs)
            argc
            (cffi:foreign-slot-value params 'DISPPARAMS 'cNamedArgs)
            0
            (cffi:mem-aref arg-err :unsigned-int) 0)
      (dotimes (i (cffi:foreign-type-size 'EXCEPINFO))
        (setf (cffi:mem-aref excep-info :unsigned-char i) 0))
      (let ((result (alloc-variant)))
        (invoke-succeeded
         (%dispatch-invoke dispatch
                           disp-id
                           IID_NULL
                           *lcid*
                           (logior DISPATCH_METHOD DISPATCH_PROPERTYGET)
                           params
                           result
                           excep-info
                           arg-err))
        (loop for i from 0 below argc
           for v = (cffi:mem-aref v-args 'VARIANT i)
           do (VariantClear v))
        result))))

(defun dispatch-get-type-info-count (dispatch)
  (cffi:with-foreign-object (count 'UINT)
    (succeeded (%dispatch-get-type-info-count dispatch count))
    (cffi:mem-aref count 'UINT)))


(defun dispatch-get-type-info (dispatch &optional (iTInfo 0))
  (cffi:with-foreign-object (pp-type-info :pointer)
    (succeeded (%dispatch-get-type-info dispatch
                                        iTInfo
                                        *lcid*
                                        pp-type-info))
    (cffi:mem-aref pp-type-info :pointer)))
