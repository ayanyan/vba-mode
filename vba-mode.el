;; vba-mode.el --- a mode for editing Excel VBA scripts

;;; Copyright (C) 2008 Yoshihiko Kakutani

;;; Author: Yoshihiko Kakutani <yoshihiko.kakutani@gmail.com>

;;; Copyright Notice:

;; This file is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation; either version 2, or (at your option)
;; any later version.
;;
;; This file is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.
;;
;; You should have received a copy of the GNU General Public License
;; along with GNU Emacs; see the file COPYING.  If not, write to
;; the Free Software Foundation, Inc., 59 Temple Place - Suite 330,
;; Boston, MA 02111-1307, USA.

;;; Commentary:

;; A major mode for editing Excel VBA codes.  Indentation and
;; colorization are supported.  Some structures can be inserted with
;; simple commands.  Completion is available for class names and
;; function names.
;;
;; The following code is a typiacl example of .emacs.
;;
;; (autoload 'vba-mode "vba-mode" "Turn a mode for VBA on." t nil)
;; (setq auto-mode-alist
;;       (cons '("\\.vba?$" . vba-mode) auto-mode-alist))

;;; Acknowledgments:

;; The author is inspired to write this code by `visual-basic-mode'
;; provided by Fred White.  The author would be grateful to him.

;;; Code:

(defvar vba-mode-indent 2)

(defvar vba-use-font-lock t)

(defvar vba-capitalize-keywords t)

(defconst vba-keywords
  '("Function" "End Function"
    "Sub" "End Sub"
    "Type" "End Type"
    "Option" "Set" "As"
    "If" "Then" "Else" "ElseIf" "End If"
    "Do" "While" "Until" "Loop" "Wend"
    "For" "Each" "In" "To" "Step" "Next"
    "Select" "Case" "Case Is" "Case Else" "End Select"
    "With" "End With"
    "Exit" "GoTo" "On Error" "End"))

(defconst vba-modifiers
  '("Dim" "Public" "Private" "Static"))

(defconst vba-operators
  '("Array" "True" "False" "Not" "And" "Or"
    "Is" "Like" "Nothing" "Null" "EOF"))

(defconst vba-types
  '("Boolean" "Integer" "Long" "Single" "Double" "String"
    "Date" "Currency" "Variant" "Object"
    "Range" "Worksheet" "Workbook"))

(defconst vba-functions
  '("IIf" "Switch" "UBound" "LBound" "IsNull" "IsNumeric" "IsEmpty"
    "Round" "Fix" "Sqr" "Rnd" "Format"
    "Chr" "ChrW" "Len" "Split" "Join" "Replace" "StrComp"
    "LCase" "UCase" "StrConv" "StrReverse"
    "Date" "Year" "Month" "Day" "Time" "Hour" "Minute" "Second"
    "Worksheets" "Activate" "Move" "Copy" "Delete" "Name"
    "Worksheets.Add" "Worksheets.Count"
    "Range" "Cells" "Rows" "Columns" "Offset" "SpecialCells"
    "Row" "Column" "Address" "Value" "Formula" "FormulaR1C1"
    "End" "CurrentRegion" "Selection" "MergeCells" "MergeArea"
    "Find" "FindNext" "Replace" "AutoFilter" "AdvancedFilter"
    "Clear" "Copy" "Cut" "Paste" "Insert" "Delete" "Sort"
    "Activate" "Select"
    "Font" "Interior" "Color" "RGB"
    "MsgBox" "InputBox"
    "vbYes" "vbNo" "vbCancel" "vbYesNo" "vbYesNoCancel"
    "xlYes" "xlNo" "xlDown" "xlTop" "xlToRight" "xlToLeft"
    "xlLastCell" "xlCellTypeVisible"))

(defconst vba-excel-functions
  '("Application.WorksheetFunction."
    "IF" "MIN" "MAX" "SMALL" "LARGE" "RANK"
    "SUM" "SUMIF" "SUMPRODUCT" "SUBTOTAL" "PRODUCT" "AVERAGE"
    "ROUND" "CEILING" "TRANC" "TEXT"
    "LEN" "CONCATENATE" "SUBSTITUTE"
    "ROW" "COL" "ISNA"
    "COUNTA" "COUNTBLANK" "COUNTIF"
    "VLOOKUP" "HLOOKUP" "MATCH" "INDEX"))

(defconst vba-defun-start-regexp
  (concat "^[ \t]*\\(Public\\|Private\\)?"
          "[ \t]*\\(Sub\\|Function\\|Type\\)\\b"))
(defconst vba-defun-end-regexp
  "^[ \t]*End \\(Sub\\|Function\\|Type\\)\\b")
(defconst vba-if-regexp "^[ \t]*#?If\\b")
(defconst vba-else-regexp "^[ \t]*#?Else[ \t]*\\(If\\)?\\b")
(defconst vba-endif-regexp "[ \t]*#?End If\\b")
(defconst vba-do-regexp "^[ \t]*Do\\b")
(defconst vba-loop-regexp "^[ \t]*Loop\\b")
(defconst vba-for-regexp "^[ \t]*For\\b")
(defconst vba-next-regexp "^[ \t]*Next\\b")
(defconst vba-select-regexp "^[ \t]*Select[ \t]+Case\\b")
(defconst vba-case-regexp "^[ \t]*Case\\b")
(defconst vba-endselect-regexp "^[ \t]*End Select\\b")
(defconst vba-with-regexp "^[ \t]*With\\b")
(defconst vba-endwith-regexp "^[ \t]*End With\\b")
(defconst vba-label-regexp "^[ \t]*\\(\\w+\\):[ \t]*$")
(defconst vba-continued-line-regexp "^.*_[ \t]*$")
(defconst vba-blank-line-regexp "^[ \t]*$")
(defconst vba-comment-regexp "^[ \t]*\\s<.*$")

(defvar vba-mode-syntax-table nil)

(unless vba-mode-syntax-table
  (setq vba-mode-syntax-table (make-syntax-table))
  (modify-syntax-entry ?\" "\"" vba-mode-syntax-table)
  (modify-syntax-entry ?\' "<" vba-mode-syntax-table)
  (modify-syntax-entry ?\n ">" vba-mode-syntax-table)
  (modify-syntax-entry ?\  " " vba-mode-syntax-table)
  (modify-syntax-entry ?\t " " vba-mode-syntax-table)
  (modify-syntax-entry ?\\ "w" vba-mode-syntax-table)
  (modify-syntax-entry ?. "." vba-mode-syntax-table)
  (modify-syntax-entry ?? "_" vba-mode-syntax-table)
  (modify-syntax-entry ?, "." vba-mode-syntax-table)
  (modify-syntax-entry ?_ "w" vba-mode-syntax-table)
  (modify-syntax-entry ?\( "()" vba-mode-syntax-table)
  (modify-syntax-entry ?\) ")(" vba-mode-syntax-table))

(defvar vba-mode-map nil "Keymap for VB mode.")

(unless vba-mode-map
  (setq vba-mode-map (make-sparse-keymap))
  (define-key vba-mode-map "\C-i" 'vba-indent-line)
  (define-key vba-mode-map "\C-c\C-i" 'vba-indent-region)
  (define-key vba-mode-map "\C-j" 'vba-newline-and-indent)
  (define-key vba-mode-map "\M-\C-q" 'vba-fill-or-indent)
  (define-key vba-mode-map "\C-ci" 'vba-insert-statement)
  (define-key vba-mode-map "\C-ce" 'vba-close-statement)
  (define-key vba-mode-map "\C-cf" 'vba-insert-function)
  (define-key vba-mode-map "\C-cx" 'vba-insert-excel-function)
  (define-key vba-mode-map "\M-i" 'expand-abbrev))

(defvar vba-mode-abbrev-table nil)

(defun vba-update-abbrev-table ()
  (interactive)
  (let* ((words
          (reverse (append vba-keywords vba-modifiers vba-operators
                           vba-types vba-functions)))
         (make-abbrev-spec
          (lambda (word) (list (downcase word) word))))
    (define-abbrev-table
      'vba-mode-abbrev-table
      (mapcar make-abbrev-spec words))))

(unless vba-mode-abbrev-table
  (vba-update-abbrev-table))

(defun vba-mode-variables ()
  (make-local-variable 'indent-line-function)
  (setq indent-line-function 'vba-indent-line)
  (make-local-variable 'comment-start)
  (setq comment-start "' ")
  (make-local-variable 'comment-start-skip)
  (setq comment-start-skip "'+ *")
  (make-local-variable 'comment-end)
  (setq comment-end "")
  (make-local-variable 'comment-end-skip)
  (setq comment-end-skip nil))

(defvar vba-mode-hook nil)

(defun vba-word-list-regexp (keys)
  (concat "\\b\\(" (mapconcat (lambda (x) x) keys "\\|") "\\)\\b"))

(defvar vba-font-lock-keywords
  (let ((vba-keyword-regexp (vba-word-list-regexp vba-keywords))
        (vba-modifier-regexp (vba-word-list-regexp vba-modifiers))
        (vba-operator-regexp (vba-word-list-regexp vba-operators)))
    (list
     (cons vba-keyword-regexp
           '((1 font-lock-keyword-face)))
     (cons (concat "^[ \t]*" vba-modifier-regexp)
           '((1 font-lock-type-face)))
     (cons "\\bAs[ \t]+\\(\\w+\\)"
           '((1 font-lock-type-face)))
     (cons (concat "^[ \t]*" vba-modifier-regexp "[ \t]+\\(\\w+\\)[ \t]*=")
           '((2 font-lock-variable-name-face)))
     (cons (concat vba-defun-start-regexp "[ \t]+\\(\\w+\\)")
           '((3 font-lock-function-name-face)))
     (cons vba-operator-regexp
           '((1 font-lock-builtin-face)))
     (cons vba-label-regexp
           '((1 font-lock-constant-face t))))))

(defun vba-enable-font-lock ()
  (make-local-variable 'font-lock-defaults)
  (setq font-lock-defaults '(vba-font-lock-keywords nil t nil)))

(defun vba-mode ()
  "A major mode for editing Excel VBA scripts.
Automatic indentation and fontification are provided.

Commands:
\\{vba-mode-map}"
  (interactive)
  (kill-all-local-variables)
  (setq major-mode 'vba-mode)
  (setq mode-name "VBA")
  (use-local-map vba-mode-map)
  (set-syntax-table vba-mode-syntax-table)
  (vba-mode-variables)
  (setq local-abbrev-table vba-mode-abbrev-table)
  (if vba-capitalize-keywords (abbrev-mode 1))
  (if vba-use-font-lock (vba-enable-font-lock))
  (run-hooks 'vba-mode-hook))

;; functions

(if (not (fboundp 'looking-back))
    (defun looking-back (regexp &optional limit greedy) t))

(defun vba-indent-line ()
  "Indent the current line."
  (interactive)
  (vba-indent-to (vba-calculate-indent)))

(defun vba-indent-to (col)
  (let* ((bol (point-at-bol))
         (top (+ bol col)))
    (save-excursion
      (beginning-of-line)
      (back-to-indentation)
      (when (/= (point) top)
        (delete-region bol (point))
        (indent-to col)))
    (when (< (point) top)
      (back-to-indentation))))

(defun vba-newline-and-indent (&optional count)
  "Insert a newline updating indentation."
  (interactive)
  (expand-abbrev)
  (vba-indent-line)
  (newline-and-indent))

(defun vba-indent-region (beg end)
  "Indent the region."
  (interactive "r")
  (if (< end beg)
      (let ((tmp beg))
        (setq beg end)
        (setq end tmp)))
  (save-excursion
    (save-restriction
      (goto-char beg)
      (beginning-of-line)
      (setq beg (point))
      (goto-char end)
      (end-of-line)
      (setq end (point))
      (narrow-to-region beg end)
      (goto-char (point-min))
      (while (< (point) (point-max))
        (if (not (looking-at vba-blank-line-regexp))
            (vba-indent-line))
        (forward-line 1)))))

(defun vba-fill-or-indent ()
  "Fill comment lines, or indent the current definition."
  (interactive)
  (save-excursion
    (let ((case-fold-search t))
      (beginning-of-line)
      (if (looking-at vba-comment-regexp)
          (fill-pragraph)
        (vba-indent-defun)))))

(defun vba-indent-defun ()
  (interactive)
  (save-excursion
    (beginning-of-line)
    (vba-end-of-defun)
    (let ((end (point)))
      (vba-beginning-of-defun)
      (vba-indent-region (point) end))))

(defun vba-beginning-of-defun ()
  (interactive)
  (let ((case-fold-search t))
    (re-search-backward vba-defun-start-regexp)))

(defun vba-end-of-defun ()
  (interactive)
  (let ((case-fold-search t))
    (re-search-forward vba-defun-end-regexp)))

(defun vba-insert-statement (statement)
  (interactive
   (list
    (completing-read
    "Structure: "
    '(("if") ("while") ("until") ("each") ("for") ("select") ("with")
      ("defun") ("sub") ("var"))
    nil t)))
  (let ((pos (point)))
    (save-excursion
      (end-of-line)
      (when (= (point) (point-max))
        (insert "\n")
        (forward-char -1))
      (unless (looking-back "^[ \t]*")
        (insert "\n"))
      (let ((beg (point))
            (word nil)
            (case-fold-search nil))
        (cond
         ((equal statement "if")
          (insert "If  Then\nElse\nEnd If")
          (setq word "If "))
         ((equal statement "while")
          (insert "Do While \nLoop")
          (setq word "While "))
         ((equal statement "until")
          (insert "Do Until \nLoop")
          (setq word "Until "))
         ((equal statement "each")
          (insert "For Each  In \nNext")
          (setq word "Each "))
         ((equal statement "for")
          (insert "For i = 1 To \nNext")
          (setq word "To "))
         ((equal statement "select")
          (insert "Select Case \nCase \nEnd Select")
          (setq word "Case "))
         ((equal statement "with")
          (insert "With \nEnd With")
          (setq word "With "))
         ((equal statement "defun")
          (insert "Public Function () As \n\nEnd Function")
          (setq word "Function "))
         ((equal statement "sub")
          (insert "Public Sub ()\n\nEnd Sub")
          (setq word "Sub "))
         ((equal statement "var")
          (insert "Dim  As ")
          (setq word "Dim ")))
        (vba-indent-region beg (point))
        (goto-char beg)
        (if word (re-search-forward word)))
      (setq pos (point)))
    (goto-char pos)))

(defun vba-close-statement ()
  (interactive)
  (let ((pos (point))
        (case-fold-search t))
    (save-excursion
      (beginning-of-line)
      (if (looking-at vba-blank-line-regexp)
          (end-of-line)
        (end-of-line)
        (insert "\n"))
      (let ((max nil)
            (i (save-excursion
                 (if (vba-find-matching-statement vba-if-regexp vba-endif-regexp)
                     (point) -1)))
            (l (save-excursion
                 (if (vba-find-matching-statement vba-do-regexp vba-loop-regexp)
                     (point) -1)))
            (f (save-excursion
                 (if (vba-find-matching-statement vba-for-regexp vba-next-regexp)
                     (point) -1)))
            (s (save-excursion
                 (if (vba-find-matching-statement vba-select-regexp vba-endselect-regexp)
                     (point) -1)))
            (w (save-excursion
                 (if (vba-find-matching-statement vba-with-regexp vba-endwith-regexp)
                     (point) -1))))
        (setq max (max i l f s w))
        (cond
         ((= max -1))
         ((= max i)
          (insert "End If"))
         ((= max l)
          (insert "Loop"))
         ((= max f)
          (insert "Next"))
         ((= max s)
          (insert "End Select"))
         ((= max w)
          (insert "End With")))
        (vba-indent-line)
        (setq pos (point))))
    (goto-char pos)))

(defun vba-insert-with-completions (words)
  (let ((completion-ignore-case t)
        (alist (mapcar (lambda (word) (list word)) words)))
    (insert
     (completing-read "Function: " alist))))

(defun vba-insert-function ()
  (interactive)
  (vba-insert-with-completions vba-functions))

(defun vba-insert-excel-function ()
  (interactive)
  (vba-insert-with-completions vba-excel-functions))

(defun vba-previous-line-of-code ()
  (let ((case-fold-search t))
    (if (not (bobp))
        (forward-line -1))
    (while (and (not (bobp))
                (or (looking-at vba-blank-line-regexp)
                    (looking-at vba-comment-regexp)))
      (forward-line -1))))

(defun vba-continuation-p ()
  (save-excursion
    (let ((case-fold-search t))
      (vba-previous-line-of-code)
      (looking-at vba-continued-line-regexp))))

(defun vba-find-original-statement ()
  (beginning-of-line)
  (when (vba-continuation-p)
    (vba-previous-line-of-code)
    (vba-find-original-statement)))

(defun vba-find-matching-statement (open-regexp close-regexp)
  (let ((level 1)
        (case-fold-search t))
    (while (> level 0)
      (vba-previous-line-of-code)
      (vba-find-original-statement)
      (cond ((looking-at close-regexp)
             (setq level (+ level 1)))
            ((looking-at open-regexp)
             (setq level (- level 1)))
            ((bobp)
             (setq level -1))))
    (= level 0)))

(defun vba-calculate-indent ()
  (save-excursion
    (save-restriction
      (widen)
      (let ((case-fold-search t))
        (beginning-of-line)
        (cond
         ((bobp)
          0)
         ((or (looking-at vba-defun-start-regexp)
              (looking-at vba-label-regexp)
              (looking-at vba-defun-end-regexp))
          0)
         ((or (looking-at vba-else-regexp)
              (looking-at vba-endif-regexp))
          (vba-find-matching-statement vba-if-regexp vba-endif-regexp)
          (current-indentation))
         ((looking-at vba-next-regexp)
          (vba-find-matching-statement vba-for-regexp vba-next-regexp)
          (current-indentation))
         ((looking-at vba-loop-regexp)
          (vba-find-matching-statement vba-do-regexp vba-loop-regexp)
          (current-indentation))
         ((looking-at vba-endwith-regexp)
          (vba-find-matching-statement vba-with-regexp vba-endwith-regexp)
          (current-indentation))
         ((looking-at vba-endselect-regexp)
          (vba-find-matching-statement vba-select-regexp vba-endselect-regexp)
          (current-indentation))
         ((looking-at vba-case-regexp)
          (vba-find-matching-statement vba-select-regexp vba-endselect-regexp)
          (+ (current-indentation) vba-mode-indent))
         ((vba-continuation-p)
          (vba-previous-line-of-code)
          (let ((bol (point-at-bol)))
            (end-of-line)
            (condition-case nil
                (up-list -1)
              (error (goto-char (point-min))))
            (if (not (bobp))
                (+ (current-column) 1)
              (goto-char bol)
              (vba-find-original-statement)
              (+ (current-indentation) 1))))
         (t
          (vba-previous-line-of-code)
          (while (and (not (bobp)) (looking-at vba-label-regexp))
            (vba-previous-line-of-code))
          (vba-find-original-statement)
          (cond
           ((or (looking-at vba-with-regexp))
            (current-indentation))
           ((or (looking-at vba-defun-start-regexp)
                (looking-at vba-if-regexp)
                (looking-at vba-else-regexp)
                (looking-at vba-do-regexp)
                (looking-at vba-for-regexp)
                (looking-at vba-select-regexp)
                (looking-at vba-case-regexp))
            (+ (current-indentation) vba-mode-indent))
           (t
            (current-indentation)))))))))

(provide 'vba-mode)

;;; vba-mode.el ends here
