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

(setf *poifile* (car ext::*command-line-argument-list*))
(setf *outfile* (cadr ext::*command-line-argument-list*))
(setf *infile* (caddr ext::*command-line-argument-list*))

;; if the CLI doesn't set the variables -- check vars.lisp
(if (eq *poifile* nil)
    (handler-case (load "vars.lisp")
      (file-error (c)
	(declare (ignore c))
	(format t "The vars.lisp file wasn't found :: variables not set"))))

;; load custom template objects
(setf *custom-cells* nil)
(handler-case (load "custom.lisp")
  (file-error (c)
    (declare (ignore c))
    (format t "Did not load custom.lisp")))

(defun init-classpath (&optional (poi-directory *poifile*))
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
    :documentation "Color for text in a cell")
   (foreground
    :initarg :foreground
    :initform nil
    :accessor foreground
    :documentation "Color for the cell foreground")
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
   (formula
    :initarg :formula
    :initform nil
    :accessor formula
    :documentation "Specifies adjustments to the column and row parameters")
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

(setf *letter-lookup* (vector #\A #\B #\C #\D #\E #\F #\G #\H #\I #\J #\K
			      #\L #\M #\N #\O #\P #\Q #\R #\S #\T #\U #\V
			      #\W #\X #\Y #\Z))

(defun column-to-excel-convert (column)
  (if (< column (length *letter-lookup*))
      (cons (elt *letter-lookup* column) nil)
      (let ((modulus (mod column (length *letter-lookup*)))
	    (quotient (floor (/ column (length *letter-lookup*)))))
	(cons (column-to-excel-convert (- quotient 1))
	      (cons (elt *letter-lookup* modulus) nil)))))

(defun flatten (l)
  (cond ((null l) nil)
        ((atom (car l)) (cons (car l) (flatten (cdr l))))
        (t (append (flatten (car l)) (flatten (cdr l))))))

(defun column-to-excel (column)
  (concatenate 'string (flatten (column-to-excel-convert column))))

(defun excel-identifier (column row)
  (concatenate 'string (column-to-excel column) (write-to-string (+ row 1))))

(defun lookup-color (color)
  (java:jcall "getIndex" (java:jfield "org.apache.poi.ss.usermodel.IndexedColors" (string color))))

(setf *cell-styles*
      '((bold . (lambda (style font cell) (java:jcall "setBold" font java:+true+)))
	(italic . (lambda (style font cell) (java:jcall "setItalic" font java:+true+)))
	(strikeout . (lambda (style font cell) (java:jcall "setStrikeout" font java:+true+)))
	(bottom . (lambda (style font cell) (java:jcall "setBorderBottom" style (java:jfield "org.apache.poi.ss.usermodel.BorderStyle" "MEDIUM"))))
	(top . (lambda (style font cell) (java:jcall "setBorderTop" style (java:jfield "org.apache.poi.ss.usermodel.BorderStyle" "MEDIUM"))))
	(left . (lambda (style font cell) (java:jcall "setBorderLeft" style (java:jfield "org.apache.poi.ss.usermodel.BorderStyle" "MEDIUM"))))
	(right . (lambda (style font cell) (java:jcall "setBorderRight" style (java:jfield "org.apache.poi.ss.usermodel.BorderStyle" "MEDIUM"))))
	(color . (lambda (style font cell) (java:jcall (java:jmethod "org.apache.poi.xssf.usermodel.XSSFFont" "setColor" (java:jclass "short")) font (lookup-color (slot-value cell 'color)))))
	(foreground . (lambda (style font cell) (progn
						  (java:jcall (java:jmethod "org.apache.poi.xssf.usermodel.XSSFCellStyle" "setFillForegroundColor" (java:jclass "short")) style (lookup-color (slot-value cell 'foreground)))
						  (java:jcall "setFillPattern" style (java:jfield "org.apache.poi.ss.usermodel.FillPatternType" "SOLID_FOREGROUND")))))))

(defun do-style (style font cell style-name)
  (if (slot-value cell style-name)
      (progn
	(funcall (cdr (assoc style-name *cell-styles*)) style font cell)
	1)
      0))

(defun build-formula (cell colnum rownum)
  (let* ((cell-internal (slot-value cell 'cell-internal))
	 (formula-format (slot-value cell 'cell-value))
	 (adjustments
	  (loop for adjustment in (slot-value cell 'formula)
	       for adj = (excel-identifier (+ colnum (car adjustment)) (+ rownum (cadr adjustment)))
	     collect adj)))
    (java:jcall "setCellFormula" (slot-value cell 'cell-internal)
		(apply #'format (cons nil (cons (slot-value cell 'cell-value) adjustments))))))

(defun build-cell (workbook colnum rownum cell)
  (if (slot-value cell 'custom)
	(build-cell workbook colnum rownum (eval (cons 'make-instance (cons ''cell (cons :cell-internal (cons (slot-value cell 'cell-internal) (cdr (assoc (slot-value cell 'custom) *custom-cells*))))))))
	(let* ((cell-internal (slot-value cell 'cell-internal))
	       (style (java:jcall "createCellStyle" workbook))
	       (font (java:jcall "createFont" workbook))
	       (style-fn (lambda (style-name) (do-style style font cell style-name))))
	  (if (slot-value cell 'formula)
	      (build-formula cell colnum rownum)
	      (java:jcall "setCellValue" cell-internal (slot-value cell 'cell-value)))
	  (if (> (apply #'+ (loop for setting in (loop for i in *cell-styles* for x = (car i) collect x)
		       for res = (funcall style-fn setting)
		       collect res)) 0)
	      (progn (java:jcall "setFont" style font)
		     (java:jcall "setCellStyle" cell-internal style))))))
	
(defun build-row (workbook sheet row rownum)
  (let* ((row-internal (java:jcall "createRow" (slot-value sheet 'sheet-internal) rownum))
	 (row-object (eval (cons 'make-instance (cons ''row (cons :row-internal (cons row-internal row)))))))
    (progn
      (loop for elt in (slot-value row-object 'row-values)
	 for colnum from 0
	 do (let ((cell (java:jcall "createCell" row-internal colnum)))
	      (cond ((stringp elt) (java:jcall "setCellValue" cell elt))
		    ((numberp elt) (java:jcall "setCellValue" cell elt))
		    ((listp elt)
		     (if elt (build-cell workbook colnum rownum (eval (cons 'make-instance (cons ''cell (cons :cell-internal (cons cell elt)))))))))))
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
      (java:jcall "setForceFormulaRecalculation" workbook java:+true+)
      (java:jcall "write" workbook file-output-stream)
      (java:jcall "close" file-output-stream))))

(init-classpath)
(delete-file *outfile*)
(worksheet *outfile* (read (open *infile* :if-does-not-exist nil)))
