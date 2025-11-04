let dadosPlanilha = []; 

function gerarCertificado() {
    const nomeAluno = document.getElementById('nome-aluno').value;
    const nomeCurso = document.getElementById('nome-curso').value;
    const cargaHoraria = document.getElementById('carga-horaria').value;
    const nomeCoordenador = document.getElementById('nome-coordenador').value;
    const corBorda = document.getElementById('cor-borda').value;
    const novoTitulo = document.getElementById('titulo-certificado').value;
    const novaDescricao = document.getElementById('descricao-curso').value;

    document.getElementById('campo-titulo').textContent = novoTitulo.toUpperCase() || 'CERTIFICADO DE MÉRITO';
    document.getElementById('campo-descricao').textContent = novaDescricao || 'pela conclusão bem-sucedida do curso de';
    document.getElementById('campo-aluno').textContent = nomeAluno.toUpperCase() || 'NOME DO ALUNO';
    document.getElementById('campo-curso').textContent = nomeCurso.toUpperCase() || 'NOME DO CURSO';

    document.getElementById('campo-carga-horaria').textContent = cargaHoraria || '8 HORAS';
    document.getElementById('campo-coordenador').textContent = 'Coordenador ' + (nomeCoordenador || 'CHARLES ROTRON').toUpperCase();
    
    document.getElementById('modelo-certificado').style.borderColor = corBorda;
    document.getElementById('certificado').style.display = 'block';
}

function baixarCertificado() {
    const modeloCertificado = document.getElementById('modelo-certificado');

    html2canvas(modeloCertificado, { 
        scale: 2,
        backgroundColor: null,
        useCORS: true
    }).then(canvas => {
        const imagemURL = canvas.toDataURL('image/png');
        const link = document.createElement('a');
        link.href = imagemURL;
        
        const nomeAlunoFormatado = document.getElementById('nome-aluno').value.replace(/\s/g, '_');
        link.download = 'certificado_' + nomeAlunoFormatado + '.png';

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }) 
} 

function iniciarGeracaoEmLote() {
    if (dadosPlanilha.length < 2) {
        alert("Por favor, faça o upload de uma planilha primeiro.");
        return;
    }

    const mapIndices = {
        nomeAluno: parseInt(document.getElementById('map-nome-aluno').value),
        nomeCurso: parseInt(document.getElementById('map-nome-curso').value),
        cargaHoraria: parseInt(document.getElementById('map-carga-horaria').value),
        tituloCertificado: parseInt(document.getElementById('map-titulo-certificado').value),
        nomeCoordenador: parseInt(document.getElementById('map-nome-coordenador').value),
    };
    
    if (mapIndices.nomeAluno === -1) {
        alert("O campo 'Nome do Aluno' é obrigatório para a geração em lote.");
        return;
    }

    const registros = dadosPlanilha.slice(1);
    let i = 0;
    
    alert(`Iniciando geração de ${registros.length} certificados. O download pode levar alguns segundos. Verifique as permissões do navegador.`);
    
    function processarProximo() {
        if (i < registros.length) {
            const linhaDados = registros[i];
            
            document.getElementById('nome-aluno').value = linhaDados[mapIndices.nomeAluno] || '';
            
            if (mapIndices.nomeCurso !== -1) {
                document.getElementById('nome-curso').value = linhaDados[mapIndices.nomeCurso] || document.getElementById('nome-curso').value;
            }
            if (mapIndices.cargaHoraria !== -1) {
                document.getElementById('carga-horaria').value = linhaDados[mapIndices.cargaHoraria] || document.getElementById('carga-horaria').value;
            }
            if (mapIndices.tituloCertificado !== -1) {
                document.getElementById('titulo-certificado').value = linhaDados[mapIndices.tituloCertificado] || document.getElementById('titulo-certificado').value;
            }
            if (mapIndices.nomeCoordenador !== -1) {
                document.getElementById('nome-coordenador').value = linhaDados[mapIndices.nomeCoordenador] || document.getElementById('nome-coordenador').value;
            }

            gerarCertificado(); 
            baixarCertificado();

            i++;
            setTimeout(processarProximo, 300);
        } else {
            alert("Geração em lote concluída! Os downloads sequenciais foram iniciados.");
        }
    }

    processarProximo(); 
}

document.addEventListener('DOMContentLoaded', () => {
    
    const modeloCertificado = document.getElementById('modelo-certificado'); 

    const inputCSV = document.getElementById('upload-csv');
    const mappingInterface = document.getElementById('mapping-interface');

    inputCSV.addEventListener('change', function(e) {
        dadosPlanilha = []; 
        mappingInterface.style.display = 'none';

        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (jsonSheet.length < 2) {
                alert("A planilha deve ter pelo menos um cabeçalho e uma linha de dados.");
                document.getElementById('upload-csv').value = '';
                return;
            }

            const cabecalhos = jsonSheet[0]; 
            const selects = [
                document.getElementById('map-nome-aluno'),
                document.getElementById('map-nome-curso'),
                document.getElementById('map-carga-horaria'),
                document.getElementById('map-titulo-certificado'), 
                document.getElementById('map-nome-coordenador') 
            ];

            selects.forEach(select => {
                if (!select) return; 

                select.innerHTML = ''; 
                select.innerHTML += `<option value="-1">Não Usar</option>`;
                cabecalhos.forEach((header, index) => {
                    select.innerHTML += `<option value="${index}">${header}</option>`;
                });
            });

            dadosPlanilha = jsonSheet; 
            mappingInterface.style.display = 'block';
            alert(`Planilha lida! ${dadosPlanilha.length - 1} registros encontrados.`);
        };
        
        reader.readAsArrayBuffer(file);
    });

    const inputLogo = document.getElementById('upload-logo');
    let imgLogo = document.getElementById('logo-certificado'); 

    if (!imgLogo) {
        imgLogo = document.createElement('img');
        imgLogo.id = 'logo-certificado';
        imgLogo.className = 'logo-certificado';
        imgLogo.style.display = 'none'; 
        modeloCertificado.prepend(imgLogo);
    }

    inputLogo.addEventListener('change', function(){
        const file = this.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                imgLogo.src = e.target.result;
                imgLogo.style.display = 'block'; 
            };
            reader.readAsDataURL(file);
        } else {
            imgLogo.style.display = 'none'; 
        }
    });

    const inputFundo = document.getElementById('upload-sobreposicao'); 
    
    inputFundo.addEventListener('change', function() {
        const file = this.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                modeloCertificado.style.backgroundImage = `url('${e.target.result}')`;
                modeloCertificado.style.backgroundSize = 'cover';
                modeloCertificado.style.backgroundRepeat = 'no-repeat';
                modeloCertificado.style.backgroundPosition = 'center center';
            };
            reader.readAsDataURL(file);
        } else {
            modeloCertificado.style.backgroundImage = 'url("../static/fundo_certificado.webp")'; 
        }
    });
    
    gerarCertificado();
});