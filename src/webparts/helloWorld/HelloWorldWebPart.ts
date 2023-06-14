import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { escape } from '@microsoft/sp-lodash-subset';
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import * as $ from "jquery";
import * as moment from 'moment';

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';


export interface IHelloWorldWebPartProps {
  description: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
}

//A interface ISPList contém as informações da lista do SharePoint à qual estamos nos conectando.
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Escritor: string;
  Id: string;
  Inicio: string;
  Fim: string;
  Avaliacao: number;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

 
  //private _environmentMessage: string = '';
  private livros: ISPList[] = []
  private itemID: string = ""
    public render(): void {
      
      this.domElement.innerHTML = `
<section>
  <div class="${styles.divBtn}">
    <button type="button" class="${styles.btnForm}" id="btnNovo" name="action">Novo cadastro</button>
    <div hidden id="formulario">
    <h1>Realize seu cadastro:</h1>
      <form class="">
        <div>
          <label for:"title">Título:</label>
          <input id="title" type="text" placeholder="Ex: Matéria Escura" name="title">
        </div>
        <div>
          <label for:"escritor">Escritor(a) do livro:</label>
          <input id="escritor" type="text" placeholder="Ex: Blake Crouch" name="escritor">
        </div>
        <div>
          <label for:"dataInicio">Data Inicial:</label>
          <input id="dataInicio" type="date" name="dataInicio">
        </div>
        <div>
          <label for:"dataFim">Data Final:</label>
          <input id="dataFim" type="date" name="dataFim"><br>
        </div>
        <div>
          <label for:"avaliacao">Avaliação:</label>
          <input id="avaliacao" type="number"  min="0" max="5" name="avaliacao">
        </div>
      </form>
      <div class="${styles.btnsCadastro}">
        <button type="submit" class="${styles.btnForm}" id="btnCadastrar">Cadastrar</button>
        <button type="submit" class="${styles.btnForm}" id="btnAtualizar">Atualizar</button>
        <button class="${styles.btnForm}" id="btnCancelar">Cancelar</button>
      </div>
        
    </div>
    
    <br>

  </div>
  <div id="spListContainer" />
</section>`;

this._renderListAsync();
this.mostrarFormulario();
this.salvarFormulario();
this.cancelarFormulario();
this.updateButton();

}

//btns
  private mostrarFormulario(): void {
    $( "#btnNovo" ).on( "click", function() {
      $('#formulario').show();
      $('#btnCadastrar').show();
      $('#btnNovo').hide();
      $('#btnAtualizar').hide();
  } );
 }

 private cancelarFormulario(): void {
  $( "#btnCancelar" ).on( "click", () => {
    this.limparDados();
    this.limparValidacao();
    $('#formulario').hide();
    $('#btnNovo').show();
    
} );
}

 private salvarFormulario(): void {
   $("#btnCadastrar").on( "click",  () => {this._postNewItem();});
   $('#formulario').hide();
   $('#btnNovo').show();
 }

 private deleteItem(): void {
  const deleteItem = document.querySelectorAll('.btnExcluir')
  deleteItem.forEach((button, index) =>{
    button.addEventListener('click', () =>{
      const itemID = this.livros[index].Id;
      this._deleteItem(itemID);
    })  
  })
 }
 
 private updateItem(): void {
  const updateItem = document.querySelectorAll('.btnUpdate')
  updateItem.forEach((button, index) =>{
    button.addEventListener('click', () =>{
      this.itemID = this.livros[index].Id;
      $('#formulario').show();
      $('#btnAtualizar').show();
      $('#btnNovo').hide();
      $('#btnCadastrar').hide();
      this.resgatarDados(this.livros[index])
    })  
  })
 }

 private updateButton(): void {
  $("#btnAtualizar").on( "click",  () => {this._updateItem(this.itemID);});
  $('#formulario').hide();
  $('#btnNovo').show();
}

private resgatarDados(livro: ISPList): void{
  $('input[name=title]').val(livro.Title);
  $('input[name=escritor]').val(livro.Escritor);
  $('input[name=dataInicio]').val(livro.Inicio.split('T')[0]);
  $('input[name=dataFim]').val(livro.Fim.split('T')[0]);
  $('input[name=avaliacao]').val(livro.Avaliacao);
}

private limparDados(): void{
  $('input[name=title]').val("");
  $('input[name=escritor]').val("");
  $('input[name=dataInicio]').val("");
  $('input[name=dataFim]').val("");
  $('input[name=avaliacao]').val("");
}

 private _renderListAsync(): void {
   this._getListData()
     .then((response) => {
      this.livros = response.value;
       this._renderList(response.value);
     })
     .catch(() => {console.log('render não realizado')
     });
 }
 
 //recuperar listas do SharePoint dentro da classe HelloWorldWebPart
 private _getListData(): Promise<ISPLists> {
   return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Cadastro de Livros')/items`, SPHttpClient.configurations.v1)
     .then((response: SPHttpClientResponse) => {
       return response.json();
      })
      .catch(() => {console.log('get não realizado');
      });
  }

  private validateItem(): boolean {
    let validated = true;
    const title = $('input[name=title]').val();
    const escritor = $('input[name=escritor]').val()
    const dataInicio = $('input[name=dataInicio]').val();
    const dataFim = $('input[name=dataFim]').val();
    const avaliacao = +$('input[name=avaliacao]').val();

    if(title === null || title === '') {
      const errorMessageVisible = $(".validationTitleMessage").is(":visible");
      if (errorMessageVisible === false) {
            $('#title').after("<span class='validationTitleMessage' style='color:red;'>Título obrigatório.</span>");
      }

      validated = false;
    }

    if(escritor === null || escritor === '') {
      const errorMessageVisible = $(".validationEscritorMessage").is(":visible");
      if (errorMessageVisible === false) {
            $('#escritor').after("<span class='validationEscritorMessage' style='color:red;'>Escritor(a) obrigatório.</span>");
      }

      validated = false;
    }

    if( avaliacao <0 || avaliacao > 5) {
      const errorMessageVisible = $(".validationAvaliacaoMessage").is(":visible");
      if (errorMessageVisible === false) {
            $('#avaliacao').after("<span class='validationAvaliacaoMessage' style='color:red;'>Avaliação de 0 a 5.</span>");
      }
      validated = false;
    }
    if(dataFim !== null && dataFim !== '') {
      if (dataInicio === null || dataInicio === '') {
        const errorMessageVisible = $(".validationInicioMessage").is(":visible");
        if (errorMessageVisible === false) {
            $('#dataInicio').after("<span class='validationInicioMessage' style='color:red;'>Preencher data início</span>");
        }
        validated = false;
      }else if(moment(dataInicio).isAfter(moment(dataFim), 'day')){
        const errorMessageVisible = $(".validationInicioMessage").is(":visible");
        if (errorMessageVisible === false) {
          $('#dataInicio').after("<span class='validationInicioMessage' style='color:red;'>Precisa ser antes da data final.</span>");
        }
        validated = false;
      }

      if(moment().isBefore(moment(dataFim), 'day')){
        const errorMessageVisible = $(".validationFimMessage").is(":visible");
        if (errorMessageVisible === false) {
            $('#dataFim').after("<span class='validationFimMessage' style='color:red;'>Não pode exceder a data atual.</span>");
        }
        validated = false;
      }
      
    }
    return validated;
  }

  private convertBodyValues(): string {
    const title = $('input[name=title]').val();
    const escritor = $('input[name=escritor]').val();
    const dataInicio = $('input[name=dataInicio]').val()
    const dataFim = $('input[name=dataFim]').val();
    const avaliacao = $('input[name=avaliacao]').val();

    let bodyInicio = ''
    let bodyFim = ''
    let bodyAvaliacao = ''


    if (dataInicio !== '' && dataInicio !== null) {
      bodyInicio = `, Inicio:'${dataInicio}'`
    }

    if (dataFim !== '' && dataFim !== null) {
      bodyFim = `, Fim:'${dataFim}'`
    }

    if (avaliacao !== '' && avaliacao !== null) {
      bodyAvaliacao = `, Avaliacao:'${avaliacao}'`
    }


    return `{Title:'${title}', Escritor:'${escritor}'${bodyInicio}${bodyFim}${bodyAvaliacao}}`;
  }

  private limparValidacao(): void {
    if ($(".validationTitleMessage").is(":visible")) {
      $(".validationTitleMessage").remove();
    }
    if ($(".validationEscritorMessage").is(":visible")) {
      $(".validationEscritorMessage").remove();
    }
    if ($(".validationAvaliacaoMessage").is(":visible")) {
      $(".validationAvaliacaoMessage").remove();
    }
    if ($(".validationInicioMessage").is(":visible")) {
      $(".validationInicioMessage").remove();
    }
    if ($(".validationFimMessage").is(":visible")) {
      $(".validationFimMessage").remove();
    }
  }

  //post
  private _postNewItem(): void {
    if(this.validateItem()){
      const spOptions: ISPHttpClientOptions ={
        "body": this.convertBodyValues()
      }
  
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Cadastro de Livros')/items`, SPHttpClient.configurations.v1, spOptions)
        .then((response: SPHttpClientResponse) => {
          this._renderListAsync();
          this.limparDados()
          this.limparValidacao()
        })
        .catch(() => {console.log("post não realizado");
        });
    }
    
  }

  //delete
  private _deleteItem(Id: string): void {
    const spOptions: ISPHttpClientOptions ={
      "headers": {"X-HTTP-Method": "DELETE", "IF-MATCH":"*"}
    }
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Cadastro de Livros')/items(${Id})`, SPHttpClient.configurations.v1, spOptions)
      .then((response: SPHttpClientResponse) => {
        this._renderListAsync();
      })
      .catch(() => {console.log("delete não realizado")});
  }

  //update
  private _updateItem(Id: string): void {
    if (this.validateItem()) {
      const spOptions: ISPHttpClientOptions ={
      "body": this.convertBodyValues(),
      "headers": {"X-HTTP-Method": "MERGE", "IF-MATCH":"*"}
    }

    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Cadastro de Livros')/items(${Id})`, SPHttpClient.configurations.v1, spOptions)
      .then((response: SPHttpClientResponse) => {
        this._renderListAsync();
        this.limparDados();
        this.limparValidacao()
      })
      .catch(() => {console.log("update não realizado")});
    }
    
  }


  //faz referência aos estilos CSS adicionados usando a stylesvariável e é usado para renderizar as informações da lista que serão recebidas da API REST
  private _renderList(items: ISPList[]): void {
    let tableItens = ''
    items.forEach((item: ISPList) => {
      const dataInicio = item.Inicio !== null ? (moment(item.Inicio)).format('DD/MM/YYYY') : ''
      const dataFim = item.Fim !== null ? (moment(item.Fim)).format('DD/MM/YYYY') : ''
      const avaliacao = item.Avaliacao !== null ? item.Avaliacao.toString() : ''

      tableItens += `
      <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@48,400,0,0" />
        <tr class="menu-item">
        <td><span class="ms-font-l">${item.Id}</span></td>
        <td><span class="ms-font-l">${item.Title}</span></td>
        <td><span class="ms-font-l">${item.Escritor}</span></td>
        <td><span class="ms-font-l">${dataInicio}</span></td>
        <td><span class="ms-font-l">${dataFim}</span></td>
        <td><span class="ms-font-l ">${avaliacao}</span></td>
        <td>
        <span class="${styles.cursorBtn} btnExcluir material-symbols-outlined" name="action">delete_forever</span>
        <span type="submit" class="${styles.cursorBtn} btnUpdate material-symbols-outlined" name="action">edit</span>
         
        </td>
        </tr>
      `;
    });
    
    const html: string = `
    <h1>Livros cadastrados:</h1>
      <table>
        <tr>
          <th>ID</th>
          <th>Livro</th>
          <th>Escritor(a)</th>
          <th>Início</th>
          <th>Fim</th>
          <th>Avaliação</th>
          <th></th>
        </tr>
        ${tableItens}
      </table>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;  
    this.deleteItem();
    this.updateItem();
    
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
              PropertyPaneTextField('description', {
                label: 'Description'
              }),
              PropertyPaneTextField('test', {
                label: 'Multi-line Text Field',
                multiline: true
              }),
              PropertyPaneCheckbox('test1', {
                text: 'Checkbox'
              }),
              PropertyPaneDropdown('test2', {
                label: 'Dropdown',
                options: [
                  { key: '1', text: 'One' },
                  { key: '2', text: 'Two' },
                  { key: '3', text: 'Three' },
                  { key: '4', text: 'Four' }
                ]}),
              PropertyPaneToggle('test3', {
                label: 'Toggle',
                onText: 'On',
                offText: 'Off'
              })
            ]
            }
          ]
        }
      ]
    };
  }
}
