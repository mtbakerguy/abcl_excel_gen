;; Call this with a command-line like the following:
;;        java -jar abcl-bin-1.6.0/abcl.jar --noinform -- ~/ss/poi-4.1.1/poi-4.1.1.jar xyz.xlsx sheets.lisp < buildsheet.lisp
;; the first argument after -- the path to the poi library
;; the second output file -- output file
;; the third argument after -- input file
;;
;; buildsheet.lisp is this file
;; 

(defpackage :export-tsv
  (:use :cl))

(in-package :export-tsv)

(setf poifile (car ext::*command-line-argument-list*))
(setf outfile (cadr ext::*command-line-argument-list*))
(setf infile (caddr ext::*command-line-argument-list*))

;; if the CLI doesn't set the variables -- check vars.lisp
(if (eq poifile nil)
    (handler-case (load "vars.lisp")
      (file-error (c)
	(declare (ignore c))
	(format t "The vars.lisp file wasn't found :: variables not set"))))

;; load custom template objects
(setf custom-cells nil)
(handler-case (load "custom.lisp")
  (file-error (c)
    (declare (ignore c))
    (format t "Did not load custom.lisp")))

(defun init-classpath (&optional (poi-directory poifile))
  (let ((*default-pathname-defaults* poi-directory))
    (dolist (jar-pathname (or (directory "**/*.jar")
                              (error "no jars found in ~S - expected Apache POI binary ~
                                        installation there"
                                     (merge-pathnames poi-directory))))
      (java:add-to-classpath (namestring jar-pathname)))))

(defclass sheet ()
  ((name
    :initarg :name
    :initform (error "Must supply a sheet name")
    :accessor name
    :documentation "Sheet name")
   (rows
    :initarg :rows
    :initform nil
    :accessor rows
    :documentation "Rows to create")
   (sheet-internal
    :initarg :sheet-internal
    :initform nil
    :accessor sheet-internal
    :documentation "Java object returned from createSheet")))

(defclass row ()
  ((row-internal
    :initarg :row-internal
    :initform (error "Must supply an internal row")
    :accessor row-internal
    :documentation "Java object returned from createRow")
   (hidden
    :initarg :hidden
    :initform nil
    :accessor hidden
    :documentation "Boolean value delineating a hidden row")
   (row-values
    :initarg :row-values
    :initform (error "Must have a value")
    :accessor row-values
    :documentation "List containing the row's cell values")))

(defclass cell ()
  ((cell-internal
    :initarg :cell-internal
    :initform (error "Must supply an internal cell")
    :accessor cell-internal
    :documentation "Java object return from createCell")
   (bold
    :initarg :bold
    :initform nil
    :accessor bold
    :documentation "Bold the text in a cell")
   (italic
    :initarg :italic
    :initform nil
    :accessor italic
    :documentation "Italicize the text in a cell")
   (strikeout
    :initarg :strikeout
    :initform nil
    :accessor strikeout
    :documentation "Strikeout the text in a cell")
   (color
    :initarg :color
    :initform nil
    :accessor color
    :documentation "Color for a cell")
   (bottom
    :initarg :bottom
    :initform nil
    :accessor bottom
    :documentation "Bottom border")
   (top
    :initarg :top
    :initform nil
    :accessor top
    :documentation "Top border")
   (left
    :initarg :left
    :initform nil
    :accessor left
    :documentation "Left border")
   (right
    :initarg :right
    :initform nil
    :accessor right
    :documentation "Right border")
   (custom
    :initarg :custom
    :initform nil
    :accessor custom
    :documentation "A predefined cell with style and content")
   (cell-value
    :initarg :cell-value
    :initform nil
    :accessor cell-value
    :documentation "Text for the cell")))

(defun lookup-color (color)
  (java:jcall "getIndex" (java:jfield "org.apache.poi.ss.usermodel.IndexedColors" (string color))))

(defun build-cell (workbook cell)
  (if (slot-value cell 'custom)
      (build-cell workbook (eval (cons 'make-instance (cons ''cell (cons :cell-internal (cons (slot-value cell 'cell-internal) (cdr (assoc (slot-value cell 'custom) custom-cells))))))))
      (let* ((cell-internal (slot-value cell 'cell-internal))
	     (style (java:jcall "createCellStyle" workbook))
	     (font (java:jcall "createFont" workbook)))
	(java:jcall "setCellValue" cell-internal (slot-value cell 'cell-value))
	(if (slot-value cell 'bold) (java:jcall "setBold" font java:+true+))
	(if (slot-value cell 'italic) (java:jcall "setItalic" font java:+true+))
	(if (slot-value cell 'strikeout) (java:jcall "setStrikeout" font java:+true+))
	(if (slot-value cell 'bottom) (java:jcall "setBorderBottom" style (java:jfield "org.apache.poi.ss.usermodel.BorderStyle" "MEDIUM")))
	(if (slot-value cell 'top) (java:jcall "setBorderTop" style (java:jfield "org.apache.poi.ss.usermodel.BorderStyle" "MEDIUM")))
	(if (slot-value cell 'left) (java:jcall "setBorderLeft" style (java:jfield "org.apache.poi.ss.usermodel.BorderStyle" "MEDIUM")))
	(if (slot-value cell 'right) (java:jcall "setBorderRight" style (java:jfield "org.apache.poi.ss.usermodel.BorderStyle" "MEDIUM")))
	(if (slot-value cell 'color)
	    (java:jcall (java:jmethod "org.apache.poi.xssf.usermodel.XSSFFont" "setColor" (java:jclass "short"))
			font (lookup-color (slot-value cell 'color))))
	(if (or (slot-value cell 'bold)
		(slot-value cell 'color)
		(slot-value cell 'underline)
		(slot-value cell 'italic)
		(slot-value cell 'strikeout)
		(slot-value cell 'bottom)
		(slot-value cell 'top)
		(slot-value cell 'left)
		(slot-value cell 'right))
	    (progn
	      (java:jcall "setFont" style font)
	      (java:jcall "setCellStyle" cell-internal style))))))

(defun build-row (workbook sheet row rownum)
  (let* ((row-internal (java:jcall "createRow" (slot-value sheet 'sheet-internal) rownum))
	 (row-object (eval (cons 'make-instance (cons ''row (cons :row-internal (cons row-internal row)))))))
    (progn
      (loop for elt in (slot-value row-object 'row-values)
	 for cellnum from 0
	 do (let ((cell (java:jcall "createCell" row-internal cellnum)))
	      (cond ((stringp elt) (java:jcall "setCellValue" cell elt))
		    ((numberp elt) (java:jcall "setCellValue" cell elt))
		    ((listp elt)
		     (if elt (build-cell workbook (eval (cons 'make-instance (cons ''cell (cons :cell-internal (cons cell elt)))))))))))
      (if (slot-value row-object 'hidden) (java:jcall "setZeroHeight" row-internal t)))))

(defun build-sheet (workbook sheet-ctor)
  (let* ((sheet (eval (cons 'make-instance (cons ''sheet sheet-ctor))))
	 (sheet-internal (java:jcall "createSheet" workbook (slot-value sheet 'name))))
    (setf (slot-value sheet 'sheet-internal) sheet-internal)
    (loop for row in (slot-value sheet 'rows)
       for rownum from 0
       do (build-row workbook sheet row rownum))
    sheet))

(defun worksheet (filename sheets)
  (let* ((workbook (java:jnew "org.apache.poi.xssf.usermodel.XSSFWorkbook"))
	 (file-output-stream (java:jnew "java.io.FileOutputStream" filename))
	 (sheets (loop for sheet-ctor in sheets
		    for sheet = (build-sheet workbook sheet-ctor)
		    collect sheet)))
    (progn
      (java:jcall "write" workbook file-output-stream)
      (java:jcall "close" file-output-stream))))

(init-classpath)
(delete-file outfile)
(worksheet outfile (read (open infile :if-does-not-exist nil)))
