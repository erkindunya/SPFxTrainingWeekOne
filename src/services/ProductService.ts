import ax from 'axios';
import { IProduct } from '../common/IProduct';

export default class ProductService {
    constructor(private url: string) {}

    public getProducts(): Promise<IProduct[]> {
        return ax.get(this.url).then((res) => {
            return res.data.value as IProduct[];
        });
    }

    public getProductById(prodid: number): Promise<IProduct> {
        return ax
            .get(this.url + `&$filter=ProductID eq ${prodid}`)
            .then((res) => {
                return res.data.value[0] as IProduct;
            });
    }
}
