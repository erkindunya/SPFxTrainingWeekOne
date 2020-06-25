import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape, fromPairs } from '@microsoft/sp-lodash-subset';

import styles from './ProductsWebPart.module.scss';
import * as strings from 'ProductsWebPartStrings';

import { IProduct } from '../../common/IProduct';
import ProductService from '../../services/ProductService';

import * as $ from 'jquery';
import 'datatables.net';
import { SPComponentLoader } from '@microsoft/sp-loader';
export interface IProductsWebPartProps {
    description: string;
}

const svcURL =
    'https://services.odata.org/V3/Northwind/Northwind.svc/Products?$format=json';

export default class ProductsWebPart extends BaseClientSideWebPart<
    IProductsWebPartProps
> {
    private provider: ProductService;

    public onInit(): Promise<void> {
        this.provider = new ProductService(svcURL);

        return Promise.resolve();
    }

    public render(): void {
        $(this.domElement).html(`
      <div class="${styles.products}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">PROUCTS!</span>
                <p class="${styles.subTitle}">LIST OF PRODUCTS.</p>
                <table id="output" width="100%">
                </table>
            </div>
          </div>
        </div>
      </div>`);

        this.provider
            .getProducts()
            .then((products: IProduct[]) => {
                $('#output', this.domElement).DataTable({
                    data: products,
                    columns: [
                        { title: 'ID' },
                        { title: 'Name' },
                        { title: 'Price' },
                        { title: 'Stock' },
                    ],
                });
            })
            .catch((err) => {
                $('#output', this.domElement).html(
                    `<span>Error loading data : ${err}</span>`
                );
            });
    }

    private getHTMLTable(items: IProduct[]): string {
        let html = `<table>
    <tr>
    <td> ProductID</td>
    <td> ProductName </td>
    <td>UnitPrice</td>
    <td>UnitsInStock</td>
    </tr>`;

        for (const p of items) {
            html += `  <tr>
                <td>  ${p.ProductID} </td>
                <td> ${p.ProductName} </td>
                <td> ${p.UnitPrice} </td>
                <td>  ${p.UnitsInStock}</td>
                </tr>`;
        }
        return html + '</table>';
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription,
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel,
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    }
}
