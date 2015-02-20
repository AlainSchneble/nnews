;;; nnews.el --- Microsoft Exchange / Exchange Web Services EWS back end for Gnus

;; COPYRIGHT 2015 Alain Schneble

;; This program is free software: you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.
;;  
;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.
;;  
;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;; Author: Alain Schneble <ans@member.fsf.org>
;; URL: https://github.com/AlainSchneble/nnews

;;; History:

;; Version 0.1: Initial release

;;; Commentary:

;; Microsoft Exchange / Exchange Web Services EWS back end for Gnus

;; This is an alpha release. If you have any questions, please contact
;; Alain Schneble <ans@member.fsf.org>.

;;; Code:

(require 'gnus-util)
(require 'gnus)
(require 'nnoo)
(require 'auth-source)

(eval-when-compile
  (require 'cl))


(nnoo-declare nnews)


;; Public backend variables

(defvoo nnews-hostname nil
  "The DNS resolvable hostname of the server on which the Exchange Web Service (EWS) is published, e.g. `mydomain.org'.")

(defvoo nnews-protocol "https"
  "The protocol to use to access the Exchange Web Service (EWS). Supported are `http' or `https'. Default is `https'.")

(defvoo nnews-path-absolute "/EWS/Exchange.asmx"
  "The absolute URL path on which the Exchange Web Service (EWS) is published. Default is `/EWS/Exchange.asmx'.
Path must start with a slash (/).

The effective URL is the concatenation of `nnews-protocol' \"://\" `nnews-hostname' `nnews-path-absolute'")

(defvoo nnews-username nil
  "Username to use for authentication against the Exchange Web Service (EWS).")


;; Private backend variables

(defvoo nnews--mailbox-state-cache nil
  ;; Format:
  ;; '(:FolderId "msgfolderroot"
  ;; 	      :SubfoldersSyncState nil
  ;; 	      :Subfolders (
  ;; 			     (:FolderId "[ItemId]" :DisplayName "Sent Items" :ItemsSyncState nil :ItemIdMap nil)
  ;; 			     (:FolderId "[ItemId]" :DisplayName "Outbox" :ItemsSyncState nil :ItemIdMap nil)
  ;; 			     (:FolderId "[ItemId]" :DisplayName "Junk Email" :ItemsSyncState nil :ItemIdMap nil)
  ;; 			     (:FolderId "[ItemId]" :DisplayName "Inbox" :ItemsSyncState nil :ItemIdMap nil)
  ;; 			     (:FolderId "[ItemId]" :DisplayName "Drafts" :ItemsSyncState nil :ItemIdMap nil)
  ;; 			     (:FolderId "[ItemId]" :DisplayName "Deleted Items" :ItemsSyncState nil :ItemIdMap nil)
  ;; 			     ))
  )


;; Backend life cycle

(deffoo nnews-open-server (server &optional definitions)
  (if (nnews-server-opened server)
      t
    (unless (assq 'nnews-hostname definitions)
      ;; If no hostname has been specified explicitly in the server definition,
      ;; assume the virtual server name as hostname.
      (setq definitions (append definitions (list (list 'nnews-hostname server)))))
    (nnoo-change-server 'nnews server definitions)
    ;; Create and initialize new local mailbox cache data structure
    (unless nnews--mailbox-state-cache
      (setq
       nnews--mailbox-state-cache
       '(:FolderId "msgfolderroot" :SubfoldersSyncState nil :Subfolders ())
       )
      )
    t)
  )

(deffoo nnews-server-opened (&optional server)
  (and (nnoo-current-server-p 'nnews server)
       nntp-server-buffer
       (gnus-buffer-live-p nntp-server-buffer))
  )

(deffoo nnews-close-server (&optional server)
  (when (nnoo-change-server 'nnews server nil)
    ;; ans/todo: save mailbox state
    (nnoo-close-server 'nnews server)
    t)
  )

(deffoo nnews-request-close ()
  t)

(deffoo nnews-status-message (&optional server)
  ;; ans/todo: implement status-message stack
  "Ok")


;; Group fetching

(deffoo nnews-request-list (&optional server)
  (with-current-buffer nntp-server-buffer
    (erase-buffer)
    (let ((credentials (nnews--credentials-query server nnews-protocol nnews-username)))
      (nnews--sync-folder-hierarchy-request (plist-get nnews--mailbox-state-cache :SubfoldersSyncState)))
    (let ((childFolders (plist-get nnews--mailbox-state-cache :Subfolders)))
      (while childFolders
	(let* ((childFolder (car childFolders))
	      (itemIdMap (plist-get childFolder :ItemIdMap)))
	  (insert (replace-regexp-in-string " " "-" (plist-get childFolder :DisplayName)))
	  (insert " ")
	  (insert (int-to-string (or (car (car itemIdMap)) 0)))
	  (insert " ")
	  (insert (int-to-string (or (car (car (nthcdr (1- (length itemIdMap)) itemIdMap))) 1)))
	  (insert " y\n")
	  (setq childFolders (cdr childFolders))
	  )
	)
      )
    )
  t
  )

(deffoo nnews-request-group (group &optional server fast info)
  (with-current-buffer nntp-server-buffer
    (erase-buffer)
    (let* ((childFolders (plist-get nnews--mailbox-state-cache :Subfolders)))
      (do ((childFolder (car childFolders) (car childFolders))) ((equal group (plist-get childFolder :DisplayName)))
	(setq childFolders (cdr childFolders)))
      (let* ((childFolder (car childFolders))
	    (itemIdMap (plist-get childFolder :ItemIdMap)))
	(insert "211 ")
	(insert (int-to-string (or (length itemIdMap) 0)))
	(insert " ")
	(insert (int-to-string (or (car (car (nthcdr (1- (length itemIdMap)) itemIdMap))) 1)))
	(insert " ")
	(insert (int-to-string (or (car (car itemIdMap)) 0)))
	(insert " ")
	(insert (replace-regexp-in-string " " "-" (plist-get childFolder :DisplayName)))
	(insert "\n")
	)
      )
    )
  t)

(deffoo nnews-close-group (group &optional server)
  nil)


;; Article fetching

(deffoo nnews-retrieve-headers (articles &optional group server fetch-old)
  (with-current-buffer nntp-server-buffer
    (erase-buffer)
    (let* ((childFolders (plist-get nnews--mailbox-state-cache :Subfolders)))
      (do ((childFolder (car childFolders) (car childFolders))) ((equal group (plist-get childFolder :DisplayName)))
	(setq childFolders (cdr childFolders)))
      (let* ((childFolder (car childFolders))
	     (itemIdMap (plist-get childFolder :ItemIdMap))
	     (itemIds (mapcar (lambda (articleId) (cdr (assoc articleId itemIdMap))) articles))
	     (credentials (nnews--credentials-query server nnews-protocol nnews-username))
	     (messages (nreverse (nnews--get-item-request (plist-get childFolder :FolderId) itemIds))))
	(while messages
	  (let ((message (car messages)))
	    ;; NOV format (tab separated list):
	    ;; `article-number', `subject', `from', `date', `id', `references', `chars', `lines', `xref', `extra'
	    (insert (int-to-string (car (rassoc (plist-get message :ItemId) itemIdMap))))
	    (insert "\t")
	    (insert (plist-get message :Subject))
	    (insert "\t")
	    (insert (format "%s <%s>" (plist-get message :FromName) (plist-get message :FromAddress)))
	    (insert "\t")
	    (insert (plist-get message :DateReceived))
	    (insert "\t")		    
	    (insert (plist-get message :ItemId))
	    (insert "\t\t")
	    (insert (int-to-string (plist-get message :BodyAsTextCharCount)))
	    (insert "\t")
	    (insert (int-to-string (plist-get message :BodyAsTextLineCount)))
	    (insert "\t\t\n")
	    )
	  (setq messages (cdr messages))
	  )
	)
      )
    )
  'nov)

(deffoo nnews-request-article (article &optional group server to-buffer)
  (with-current-buffer to-buffer
    (erase-buffer)
    (let* ((childFolders (plist-get nnews--mailbox-state-cache :Subfolders)))
      (do ((childFolder (car childFolders) (car childFolders))) ((equal group (plist-get childFolder :DisplayName)))
	(setq childFolders (cdr childFolders)))
      (let* ((childFolder (car childFolders))
	     (itemIdMap (plist-get childFolder :ItemIdMap))
	     (itemId (cdr (assoc article itemIdMap)))
	     (credentials (nnews--credentials-query server nnews-protocol nnews-username))
	     (messages (nnews--get-item-request (plist-get childFolder :FolderId) (list itemId))))
	(let ((message (car messages)))
	  (insert (format "From: %s <%s>" (plist-get message :FromName) (plist-get message :FromAddress)))
	  (newline)
	  (dolist (header (plist-get message :Headers))
	    (insert (car header))
	    (insert ": ")
	    (insert (cdr header))
	    (insert "\n")
	    )
	  (insert "\n")
	  (insert (plist-get message :BodyAsText))

	  (save-excursion (search-backward "application/ms-tnef" nil t) (replace-match "text/plain"))
	  (nnheader-ms-strip-cr)

	  (cons group (int-to-string (car (rassoc (plist-get message :ItemId) itemIdMap))))
	  )
	)
      )
    
    )
  )

(deffoo nnews-request-post (&optional server)
  nil)


;; Private backend

(defun nnews--credentials-query (hostname protocol username)
  (let* ((auth-source-creation-prompts
          '((user  . "EWS user at %h: ")
            (secret . "EWS password for %u@%h: ")))
         (credentials (nth 0 (auth-source-search :max 1
                                           :host hostname
                                           :port protocol
                                           :user username
                                           :require '(:user :secret)
                                           :create t))))
    (if credentials
        (list (plist-get credentials :user)
	      (let ((secret (plist-get credentials :secret)))
		(if (functionp secret)
		    (funcall secret)
		  secret))
	      (plist-get credentials :save-function))
      nil)))

(defun nnews--ewsurl ()
  (concat nnews-protocol "://" nnews-hostname nnews-path-absolute))

(defun nnews--credentials-http-basic-auth ()
  (concat "Basic " (base64-encode-string (concat (car credentials) ":" (cadr credentials)))))

(defun nnews--xml-get-node (pathToNode xmlDom)
  (let ((selectedNode xmlDom))
    (dolist (node pathToNode selectedNode)
      (setq selectedNode (assq node selectedNode)))
    )
  )

(defun nnews--xml-get-node-attr (xmlElement xmlAttributeSymbol)
  (cdr (assq xmlAttributeSymbol (nth 1 xmlElement))))


(defun nnews--count-string-matches (regexp string)
  (nnews--count-string-matches-recursive regexp string 0))

(defun nnews--count-string-matches-recursive (regexp string start)
  (if (string-match regexp string start)
      (1+ (nnews--count-string-matches-recursive regexp string (match-end 0)))
    0))

;; EWS requests

;; SyncFolderHierarchy request

(defun nnews--sync-folder-hierarchy-request (subfoldersSyncState)
  (let ((url-request-method "POST")
	(url-request-extra-headers `(("Content-Type" . "text/xml")
				     ("Authorization" . ,(nnews--credentials-http-basic-auth))))
	(url-request-data
	 (concat
	  "<?xml version=\"1.0\" encoding=\"utf-8\"?>
<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">
  <soap:Header>
    <t:RequestServerVersion Version=\"Exchange2013\" />
  </soap:Header>
  <soap:Body>
    <m:SyncFolderHierarchy>
      <m:FolderShape>
        <t:BaseShape>AllProperties</t:BaseShape>
      </m:FolderShape>
      <m:SyncFolderId>
        <t:DistinguishedFolderId Id=\"msgfolderroot\" />
      </m:SyncFolderId>"
	  (if subfoldersSyncState (concat "<m:SyncState>" subfoldersSyncState "</m:SyncState>") "")
	  "	 
    </m:SyncFolderHierarchy>
  </soap:Body>
</soap:Envelope>"
	  )
	 )
	)
    (with-current-buffer (url-retrieve-synchronously (nnews--ewsurl))
      ;;(switch-to-buffer (current-buffer))
      (let ((remoteFolderChanges (nnews-sync-folder-hierarchy-request-parser (libxml-parse-xml-region (point) (point-max)))))
	(nnews--mailbox-state-cache-sync-subfolders nnews--mailbox-state-cache remoteFolderChanges))
      )))


(defun nnews-sync-folder-hierarchy-request-parser (response)
  (pcase response
    (`(Envelope ,_ ,_ (Body ,_ (SyncFolderHierarchyResponse ,_ (ResponseMessages ,_ (SyncFolderHierarchyResponseMessage ,_ ,_ ,syncState ,_ ,changes)))))
     (let ((changeNodes (nthcdr 2 changes))
	   changeList)
       (while changeNodes
     	 (pcase (car changeNodes)
     	   (`(,change ,_ (Folder ,_ ,folderId ,_ (FolderClass ,_ "IPF.Note") ,displayName ,_ ,_ ,_ ,_))
	    (setq changeList (cons (list :FolderId (cdr (assoc 'Id (nth 1 folderId))) :Change change :DisplayName (nth 2 displayName)) changeList)))
     	   )
     	 (setq changeNodes (cdr changeNodes))
     	 )
       `(:FolderId "msgfolderroot" :SubfoldersSyncState ,(nth 2 syncState) :Subfolders ,changeList)))
    (_ 'error))
  )

(defun nnews--mailbox-state-cache-sync-subfolders (localFolder remoteFolderChanges)
  (let ((remoteSubfoldersSyncState (plist-get remoteFolderChanges :SubfoldersSyncState))
	(remoteSubfolderChanges (plist-get remoteFolderChanges :Subfolders))
	(localSubfolders (plist-get localFolder :Subfolders)))
    ;; This folder: update local `:SubfoldersSyncState'
    (plist-put localFolder :SubfoldersSyncState remoteSubfoldersSyncState)
    ;; Subfolders: sync local state based on C(R)UD changeset from SyncFolderHierarchy response
    (dolist (remoteSubfolderChange remoteSubfolderChanges)
      (case (plist-get remoteSubfolderChange :Change)
	('Create
	 (push remoteSubfolderChange localSubfolders))
	('Update
	 (let ((localFolder (find-if (lambda (entry) (eq (plist-get entry :FolderId) (plist-get remoteSubfolderChange :FolderId))) localSubfolders)))
	   (when localFolder (plist-set localFolder :DisplayName (plist-get remateFolderChange :DisplayName)))))
	('Delete
	 (delete-if (lambda (entry) (eq (plist-get entry :FolderId) (plist-get remoteSubfolderChange :FolderId))) localSubfolders))
	)
      )
    (plist-put localFolder :Subfolders localSubfolders)
    ;; Sync all subfolders
    (dolist (localSubfolder localSubfolders)
      (nnews--mailbox-state-cache-sync-subfolder-items
       localSubfolder
       (nnews--sync-item-hierarchy-request (plist-get localSubfolder :FolderId) (plist-get localSubfolder :ItemsSyncState))))
    )
  )

(defun nnews--mailbox-state-cache-sync-subfolder-items (localFolder remoteFolderItemChanges)
  (let* ((remoteItemsSyncState (plist-get remoteFolderItemChanges :ItemsSyncState))
	 (remoteItemChanges (plist-get remoteFolderItemChanges :ChildItems)))
    ;; This folder: update local `:ItemsSyncState'
    (plist-put localFolder :ItemsSyncState remoteItemsSyncState)
    ;; Items: sync local state based on C(R)UD changeset from SyncFolderItems response
    (dolist (remoteItemChange remoteItemChanges)
      (case (plist-get remoteItemChange :Change)
	('Create
	 (let ((itemId (plist-get remoteItemChange :ItemId))
	       (itemIdMap (plist-get localFolder :ItemIdMap)))
	   (unless (rassoc itemId itemIdMap)
	     (let ((articleId (1+ (or (car (car itemIdMap)) 0))))
	       (plist-put localFolder :ItemIdMap (cons `(,articleId . ,itemId) itemIdMap))
	       )
	     )
	   )
	 )
	('Update)
	('Delete)
	)
      )
    )
  )

;; SyncFolderItems request

(defun nnews--sync-item-hierarchy-request (folderId syncState)
  (let ((url-request-method "POST")
	(url-request-extra-headers `(("Content-Type" . "text/xml")
				     ("Authorization" . ,(nnews--credentials-http-basic-auth))))
	(url-request-data
	 (concat
	  "<?xml version=\"1.0\" encoding=\"utf-8\"?>
<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">
  <soap:Header>
    <t:RequestServerVersion Version=\"Exchange2013\" />
  </soap:Header>
  <soap:Body>
    <m:SyncFolderItems>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:ItemShape>
      <m:SyncFolderId>
        <t:FolderId Id=\""
	  folderId
	  "\" />
      </m:SyncFolderId>
      <m:MaxChangesReturned>512</m:MaxChangesReturned>
      <m:SyncScope>NormalItems</m:SyncScope>
    </m:SyncFolderItems>
  </soap:Body>
</soap:Envelope>"
	  )
	 )
	)
    (with-current-buffer (url-retrieve-synchronously (nnews--ewsurl))
      ;(switch-to-buffer (current-buffer))
      (nnews-sync-item-hierarchy-request-parser (libxml-parse-xml-region (point) (point-max))))
    )
  )

(defun nnews-sync-item-hierarchy-request-parser (response)
  (pcase response
    (`(Envelope ,_ ,_ (Body ,_ (SyncFolderItemsResponse ,_ (ResponseMessages ,_ (SyncFolderItemsResponseMessage ,_ ,_ ,syncState ,_ ,changes)))))
     (let ((changeNodes (nthcdr 2 changes))
	   changeList)
       (while changeNodes
     	 (pcase (car changeNodes)
     	   (`(,change ,_ (Message ,_ ,itemId))
	    (setq changeList (cons (list :ItemId (cdr (assoc 'Id (nth 1 itemId))) :Change change) changeList)))
     	   )
     	 (setq changeNodes (cdr changeNodes))
     	 )
       `(:ItemsSyncState ,(nth 2 syncState) :ChildItems ,changeList)))
    (_ 'error))
  )


;; GetItem request

(defun nnews--get-item-request (folderId itemIds)
  (let ((url-request-method "POST")
	(url-request-extra-headers `(("Content-Type" . "text/xml")
				     ("Authorization" . ,(nnews--credentials-http-basic-auth))))
	(url-request-data
	 (concat
	 "<?xml version=\"1.0\" encoding=\"utf-8\"?>
<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">
  <soap:Header>
    <t:RequestServerVersion Version=\"Exchange2013\" />
  </soap:Header>
  <soap:Body>
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI=\"item:ItemId\" />
          <t:FieldURI FieldURI=\"item:DateTimeReceived\" />
          <t:FieldURI FieldURI=\"message:From\" />
          <t:FieldURI FieldURI=\"item:Subject\" />
          <t:FieldURI FieldURI=\"item:TextBody\" />
          <t:FieldURI FieldURI=\"item:Body\" />
          <t:FieldURI FieldURI=\"item:InternetMessageHeaders\" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>"
	 (mapconcat (lambda (itemId) (concat "<t:ItemId Id=\"" itemId "\" />")) itemIds "")
	"
      </m:ItemIds>
    </m:GetItem>
  </soap:Body>
</soap:Envelope>"))
	)
    (with-current-buffer (url-retrieve-synchronously (nnews--ewsurl))
      ;;(switch-to-buffer (current-buffer)))
      (nnews--get-item-request-parser (libxml-parse-xml-region (point) (point-max))))
    )
  )

(defun nnews--get-item-request-parser (response)
  (let ((responseMessages
	 (nthcdr 2 (nnews--xml-get-node '(Body GetItemResponse ResponseMessages) response)))
	messageList)
    (while responseMessages
      (let* ((responseMessage (car responseMessages))
	     (messageNode (nnews--xml-get-node '(Items Message) responseMessage))
	     (itemId (nnews--xml-get-node-attr (nnews--xml-get-node '(ItemId) messageNode) 'Id))
	     (fromMailboxNode (nnews--xml-get-node '(Mailbox) (nnews--xml-get-node '(From) messageNode)))
	     (fromName (nth 2 (nnews--xml-get-node '(Name) fromMailboxNode)))
	     (fromAddress (nth 2 (nnews--xml-get-node '(EmailAddress) fromMailboxNode)))
	     (date (caddr (nnews--xml-get-node '(DateTimeReceived) messageNode)))
	     (subject (car (nthcdr 2 (nnews--xml-get-node '(Subject) messageNode))))
	     (body (car (nthcdr 2 (nnews--xml-get-node '(TextBody) messageNode))))
	     (lines (nnews--count-string-matches "\n" body))
	     (chars (length body))
	     (headers (nthcdr 2 (nnews--xml-get-node '(InternetMessageHeaders) messageNode)))
	     (headersParsed nil)
	     )
	(dolist (header headers)
	  (setq headersParsed (cons (cons (nnews--xml-get-node-attr header 'HeaderName) (elt header 2)) headersParsed))
	  )
	(setq messageList (cons `(:ItemId ,itemId :FromName ,fromName :FromAddress ,fromAddress :DateReceived ,date :Subject ,subject :BodyAsText ,body :BodyAsTextLineCount ,lines :BodyAsTextCharCount ,chars :Headers ,headersParsed) messageList))
	)
      (setq responseMessages (cdr responseMessages))
      )
    messageList
    )
  )


;; GNUS backend registration
;; server-marks will trigger call to request-set-mark
(unless (assoc "nnews" gnus-valid-select-methods)
  (gnus-declare-backend "nnews" 'post-mail 'address 'prompt-address 'physical-address 'respool 'server-marks))

(provide 'nnews)
