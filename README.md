# Lista-2-Ex-1
descrição : ' '
anfitrião : EXCEL
api_set : {}
script :
  content: "const field1 = document.getElementById (\" name \ ") as HTMLInputElement; \ r \ nconst field2 = document.getElementById (\" type \ ") as HTMLSelectElement; \ r \ nconst field3 = document.getElementById (\" raça \ ") como HTMLInputElement; \ r \ nconst field4 = document.getElementById (\" weight \ ") como HTMLInputElement; \ r \ nconst save = document.getElementById (\" save \ ") como HTMLButtonElement; \ r \ nconst mostrar = document.getElementById (\ "show \") como HTMLButtonElement; \ r \ n \ r \ nconst pets = []; \ r \ n \ r \ nsave.addEventListener (\ "click \", () => {\ r \ n Excel.run (assíncrono (contexto) => {\ r \ n const sheet = context.workbook.worksheets.getActiveWorksheet (); \ r \ n sheet.getUsedRange (). clear (); \ r \ n \ r \ n if (! field1.value ||! field2.value ||! field3.value ||! field4.value) {\ r \ n sheet.getRange (\ "A1 \ "). Values ​​= [[\" Todos os campos são obrigatórios! \ "]]; \ R \ n sheet.getUsedRange (). Format.autofitColumns (); \ r \ n return; \ r \ n} \ r \ n \ r \ n const pet = {\ r \ n nome: field1.value, \ r \ n type: field2.value, \ r \ n raça: field3.value, \ r \ n peso: field4.valueAsNumber \ r \ n}; \ r \ n \ r \ n pets.push (animal de estimação); \ r \ n \ r \ n field1.value = \ "\"; \ r \ n field2.value = \ "\" ; \ r \ n field3.value = \ "\"; \ r \ n field4.value = \ "\"; \ r \ n}); \ r \ n}); \ r \ n \ r \ nshow. addEventListener (\ "click \", () => {\ r \ n Excel.run (async (context) => {\ r \ n const sheet = context.workbook.worksheets.getActiveWorksheet (); \ r \ n sheet .getUsedRange (). clear (); \ r \ n \ r \ n para (let i = 1; i <= pets.length; i ++) {\ r \ n const pet = pets [i - 1]; \ r \ n sheet.getRange (`A $ {i}: D $ {i}`).valores = [[pet.name, pet.type, pet.breed, pet.weight]]; \ r \ n} \ r \ n \ r \ n sheet.getRange (\ "D1: D999 \"). numberFormat = [[\ "0.0 \"]]; \ r \ n sheet.getUsedRange (). Format.autofitColumns (); \ r \ n}); \ r \ n}); \ r \ n "
  linguagem : texto datilografado
modelo :
  conteúdo : " <h1> Cadastro de Animais de estimação </h1> \ r \ n \ r \ n <div> \ r \ n \ t <input id = \" nome \ " placeholder = \" Nome \ " > \ r \ n \ t <select id = \ " type \" > \ r \ n \ t \ t <option value = \ "\" > Tipo </option> \ r \ n \ t \ t <option> Cachorro </ option > \ r \ n \ t \ t <option> Gato </option> \ r \ n \ t </select> \ r \ n \ t <input id = \ " raça \"placeholder = \ " Raça \" > \ r \ n \ t<input id = \ " weight \" type = \ " number \" placeholder = \ " Peso \" > \ r \ n </div> \ r \ n \ r \ n <br> \ r \ n \ r \ n <div> \ r \ n \ t <button id = \ " save \" > Salvar </button> \ r \ n \ t <button id = \ " show \" > Exibir na Planilha </button> \ r \ n </div> "
  linguagem : html
estilo :
  conteúdo : ' '
  idioma : css
bibliotecas : | -
  https://appsforoffice.microsoft.com/lib/1/hosted/office.js
  @ types / office-js
  core-js@2.4.1/client/core.min.js
  @ types / core-js
