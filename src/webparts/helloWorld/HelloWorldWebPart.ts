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
  Inicio:number;
  Fim:number;
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
        <input id="title" type="text" placeholder="Título do livro" name="title"><br>
        <input id="escritor" type="text" placeholder="Escritor(a) do livro" name="escritor"><br>
        <input id="dataInicio" type="date" name="dataInicio"><br>
        <input id="dataFim" type="date" name="dataFim"><br>
        <input id="avaliacao" type="number" placeholder="Avaliação (Nota 0 a 5)" min="0" max="5" name="avaliacao">
        <br>
        
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
  $( "#btnCancelar" ).on( "click", function() {
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
    })  
  })
 }

 private updateButton(): void {
  $("#btnAtualizar").on( "click",  () => {this._updateItem(this.itemID);});
  $('#formulario').hide();
  $('#btnNovo').show();
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
   return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('CadastrodeLivros')/items`, SPHttpClient.configurations.v1)
     .then((response: SPHttpClientResponse) => {
       return response.json();
      })
      .catch(() => {console.log('get não realizado');
      });
  }

  //post
  private _postNewItem(): void {
    const title = $('input[name=title]').val();
    const escritor = $('input[name=escritor]').val();
    const spOptions: ISPHttpClientOptions ={
      "body":`{Title:'${title}', Escritor:'${escritor}' }`
    }

    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('CadastrodeLivros')/items`, SPHttpClient.configurations.v1, spOptions)
      .then((response: SPHttpClientResponse) => {
        this._renderListAsync();
      })
      .catch(() => {console.log("post não realizado");
      });
  }

  //delete
  private _deleteItem(Id: string): void {
    const spOptions: ISPHttpClientOptions ={
      "headers": {"X-HTTP-Method": "DELETE", "IF-MATCH":"*"}
    }
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('CadastrodeLivros')/items(${Id})`, SPHttpClient.configurations.v1, spOptions)
      .then((response: SPHttpClientResponse) => {
        this._renderListAsync();
      })
      .catch(() => {console.log("delete não realizado")});
  }

  //update
  private _updateItem(Id: string): void {
    const title = $('input[name=title]').val();
    const escritor = $('input[name=escritor]').val();
    const spOptions: ISPHttpClientOptions ={
      "body":`{Title:'${title}', Escritor:'${escritor}'}`,
      "headers": {"X-HTTP-Method": "MERGE", "IF-MATCH":"*"}
    }
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('CadastrodeLivros')/items(${Id})`, SPHttpClient.configurations.v1, spOptions)
      .then((response: SPHttpClientResponse) => {
        this._renderListAsync();
      })
      .catch(() => {console.log("update não realizado")});
  }


  //faz referência aos estilos CSS adicionados usando a stylesvariável e é usado para renderizar as informações da lista que serão recebidas da API REST
  private _renderList(items: ISPList[]): void {
    let tableItens = ''
    items.forEach((item: ISPList) => {
      tableItens += `
      <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@48,400,0,0" />
        <tr class="menu-item">
        <td><span class="ms-font-l">${item.Id}</span></td>
        <td><span class="ms-font-l">${item.Title}</span></td>
        <td><span class="ms-font-l">${item.Escritor}</span></td>
        <td><span class="ms-font-l">${item.Inicio}</span></td>
        <td><span class="ms-font-l">${item.Fim}</span></td>
        <td><span class="ms-font-l">${item.Avaliacao}</span></td>
        <td>
        <span class="${styles.cursorBtn} btnExcluir material-symbols-outlined" name="action">delete_forever</span>
        <span type="submit" class="${styles.cursorBtn} btnUpdate material-symbols-outlined" name="action">edit</span>
         
        </td>
        </tr>
      `;
    });
    
    const html: string = `
      <table>
        <tr>
          <th>ID</th>
          <th>Livro</th>
          <th>Escritor(a)</th>
          <th>Início</th>
          <th>Fim</th>
          <th>Avaliação</th>
          <th>Ações</th>
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
