;;; outlookedit.el --- As updated by me.

;;; http://www.emacswiki.org/emacs/MsOutlook
;;; https://github.com/dholm/outlookedit/blob/master/outlookedit.el
;;
;; Author: Various
;; Package-Version: 20150624.93625
;; Keywords: convenience

;;; Commentary:
;; use a couple of Tcl scripts to invoke my Tcl scripts to do COM
;; related actions on Outlook to get and replace text in the
;; reply/compose boxes allowing it to be edited in Emacs

(defgroup mno nil "Customization for mno" :group nil)

(defcustom
  mno-get-outlook-body
  (concat "cscript //B //Job:getMessage "
          (convert-standard-filename
           (expand-file-name
            (locate-user-emacs-file "lisp/outlook_emacs.wsf"))))
  "Command line for fetching the body of an Outlook message."
  :type 'string
  :group 'mno)

(defcustom mno-put-outlook-body
  (concat "cscript //B //Job:putMessage "
          (convert-standard-filename
           (expand-file-name
            (locate-user-emacs-file "lisp/outlook_emacs.wsf"))))
  "Command line for storing the body to an Outlook message."
  :type 'string
  :group 'mno)

(defcustom mno-outlook-default-justification
  'left
  "Default justification for Outlook messages."
  :type '(choice (const left)
                 (const full)
                 (const none)
                 (const right)
                 (const center))
  :group 'mno)

;;;###autoload
(global-set-key "\C-coe" 'mno-edit-outlook-message)
;;;###autoload
(global-set-key "\C-cos" 'mno-save-outlook-message)

;;;###autoload
(defun mno-edit-outlook-message ()
  "* Slurp an outlook message into a new buffer ready for editing

The message must be in the active Outlook window.  Typically the
user would press the Reply or Reply-all button in Outlook then
switch to Emacs and invoke \\[mno-edit-outlook-message]

Once all edits are done, the function mno-save-outlook-message
(invoked via \\[mno-save-outlook-message]) can be used to send the
newly edited text back into the Outlook window.  It is important that
the same Outlook window is current in outlook as was current when the
edit was started when this command is invoked."
  (interactive)
  (save-excursion
    (let ((buf (get-buffer-create "*Outlook Message*"))
         (body (shell-command-to-string mno-get-outlook-body)))
      (switch-to-buffer buf)
      (message-mode) ; enables automagic reflowing of text in quoted
                     ; sections
      (setq default-justification mno-outlook-default-justification)
      (setq body (replace-regexp-in-string "\r" "" body))
      (delete-region (point-min) (point-max))
      (insert body)
      (goto-char (point-min)))))

;;;###autoload
(defun mno-save-outlook-message ()
  "* Send the outlook message buffer contents back to Outlook current window

Unfortunately, Outlook 2000 then renders this text as Rich Text format
rather than plain text, overriding any user preference for plain text.
The user then needs to select Format->Plain text in the outlook
compose window to reverse this.

Outlook 2002 apparently has a BodyFormat parameter to control this."
  (interactive)
  (save-excursion
    (let ((buf (get-buffer "*Outlook Message*")))
      (set-buffer buf)
      (when (= (shell-command-on-region
                (point-min) (point-max) mno-put-outlook-body)
               0)
        (set-buffer-modified-p 'nil) ; now kill-buffer won't complain!
        (kill-buffer "*Outlook Message*")))))

(provide 'outlookedit)
;;; outlookedit.el ends here
