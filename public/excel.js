let selectedFile

document.getElementById('input').addEventListener('change', (event) => {
    selectedFile = event.target.files   [0]
})


document.getElementById('button').addEventListener('click', () => {
    if(selectedFile){
        let fileReader = new FileReader()
        fileReader.readAsBinaryString(selectedFile)
        fileReader.onload = (event) => {
            let data = event.target.result
            let workbook = XLSX.read(data, {type:'binary'})
            // console.log(workbook)
            workbook.SheetNames.forEach(sheet => {
            let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet])    
            
            const retirarVazio = item => item !== ''
            const pon = item => item.PON   
            const recurso = item => item.Recurso
            const cidade = item => item.Cidade
            const ard = item => item['ID de Armário']       
            
            const create = (pon) => {
                return pon
            }

            const getPon = rowObject.map(pon).filter(retirarVazio)
            const getCidade = rowObject.map(cidade).filter(retirarVazio)
            const getRecurso = rowObject.map(recurso).filter(retirarVazio)

            // console.log(create(getPon))

            document.getElementById('jsondata').innerHTML = JSON.stringify(rowObject, undefined, 4)
          
            });
        }
    }    
})

