(defun fix-special-chars (str)
  (vl-string-subst "ç" "Ã§" 
    (vl-string-subst "ã" "Ã£"
      (vl-string-subst "ç" "ç"
        (vl-string-subst "ã" "ã" str)))))

(defun c:criar_clash2 (/ file line x y z objects id entity1 entity2 layer1 layer2
                           filepath imgfolder idnum imagepath img_offset)
  (princ "\nIniciando o programa...")
  
  ; Verifica se o bloco "clash" existe
  (if (not (tblsearch "BLOCK" "clash"))
    (progn
      (princ "\nERRO: O bloco 'clash' não existe no desenho!")
      (exit)
    )
  )
  
  ; Solicita o arquivo TXT
  (setq filepath (getfiled "Selecione o arquivo TXT" "" "txt" 8))
  (princ (strcat "\nArquivo selecionado: " filepath))
  
  ; Solicita a pasta de imagens
  (setq imgfolder (getstring "\nDigite o caminho completo da pasta com as imagens: "))
  (princ (strcat "\nPasta de imagens: " imgfolder))
  
  (if filepath
    (progn
      (setq file (open filepath "r"))
      (if (not file)
        (progn
          (princ "\nERRO: Não foi possível abrir o arquivo!")
          (exit)
        )
      )
      
      (setvar "ATTDIA" 0)
      (setvar "ATTREQ" 1)
      
      (while (setq line (read-line file))
        (setq line (vl-string-trim " \t\r\n" (fix-special-chars line)))
        (princ (strcat "\nProcessando linha: " line))
        
        (cond 
          ((wcmatch line "X:*")
           (setq x (atof (substr line 3)))
           (princ (strcat "\nX definido: " (rtos x)))
          )
          ((wcmatch line "Y:*")
           (setq y (atof (substr line 3)))
           (princ (strcat "\nY definido: " (rtos y)))
          )
          ((wcmatch line "Z:*")
           (setq z (atof (substr line 3)))
           (princ (strcat "\nZ definido: " (rtos z)))
          )
          ((wcmatch line "Objetos:*")
           (setq objects line)
           (princ (strcat "\nObjetos: " objects))
          )
          ((wcmatch line "ID:*")
           (setq id (substr line 4))
           (princ (strcat "\nID: " id))
          )
          ((wcmatch line "Entity1:*")
           (setq entity1 line)
          )
          ((wcmatch line "Entity2:*")
           (setq entity2 line)
          )
          ((wcmatch line "Layer1:*")
           (setq layer1 line)
          )
          ((wcmatch line "Layer2:*")
           (setq layer2 line)
          )
          ((wcmatch line "-*")
           (if (and x y z)
             (progn
               (princ (strcat "\nTentando inserir bloco em: X=" (rtos x) " Y=" (rtos y) " Z=" (rtos z)))
               
               ; Tenta inserir o bloco
               (command "._insert" "clash" (list x y z) "1" "1" "0"
                        objects 
                        (if id id " ")
                        (if entity1 entity1 " ")
                        (if (= entity2 "") " " entity2)
                        (if layer1 layer1 " ")
                        (if (= layer2 "") " " layer2)
               )
               
               ; Se tiver pasta de imagens, tenta inserir a imagem
               (if (and imgfolder id)
                 (progn
                   (setq imagepath (strcat imgfolder "\\" (vl-string-trim " " id) ".jpg"))
                   (princ (strcat "\nProcurando imagem: " imagepath))
                   
                   (if (findfile imagepath)
                     (progn
                       (setq img_offset 10)
                       (princ "\nInserindo imagem...")
                       (command "._IMAGEATTACH" imagepath (list (+ x img_offset) y z) "1" "1" "0")
                     )
                     (princ (strcat "\nImagem não encontrada: " imagepath))
                   )
                 )
               )
               
               (setq x nil y nil z nil)
               (princ "\n---Bloco processado---")
             )
             (princ "\nERRO: Coordenadas incompletas!")
           )
          )
        )
      )
      
      (close file)
      (setvar "ATTDIA" 1)
      (command "._ZOOM" "E")
      (princ "\nProcessamento concluído!")
    )
    (princ "\nNenhum arquivo selecionado.")
  )
  (princ)
)