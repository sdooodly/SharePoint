import * as React from "react";
import { IFluentUiDetailsListWithCustomItemColumnProps } from "./IFluentUiDetailsListWithCustomItemColumnProps";
import { Web } from "@pnp/sp/presets/all";
import {
  IColumn,
  DetailsList,
  SelectionMode,
  DetailsListLayoutMode,
  mergeStyles,
  Link,
  Image,
  ImageFit,
} from "@fluentui/react";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface ICustomColumnDetailsListStates {
  Items: IDocument[];
  columns: any;
  isColumnReorderEnabled: boolean;
}

export interface IDocument {
  CustomerName: string;
  CustomerEmail: string;
  ProductName: string;
  OrderDate: any;
  ProductDescription: any;
  ProductImage: string;
}

export default class FluentUiDetailsListWithCustomItemColumn extends React.Component<
  IFluentUiDetailsListWithCustomItemColumnProps,
  ICustomColumnDetailsListStates
> {
  constructor(props) {
    super(props);
    const columns: IColumn[] = [
      {
        key: "ProductImage",
        name: "",
        fieldName: "ProductImage",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "ProductName",
        name: "Product Name",
        fieldName: "ProductName",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "CustomerName",
        name: "Customer Name",
        fieldName: "CustomerName",
        minWidth: 70,
        maxWidth: 90,
        isRowHeader: true,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "CustomerEmail",
        name: "Customer Email",
        fieldName: "CustomerEmail",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "number",
        isPadded: true,
      },

      {
        key: "OrderDate",
        name: "Order Date",
        fieldName: "OrderDate",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
      },
      {
        key: "ProductDescription",
        name: "Product Description",
        fieldName: "ProductDescription",
        minWidth: 210,
        maxWidth: 350,
        isResizable: true,
        data: "string",
      },
    ];

    this.state = {
      Items: [],
      columns: columns,
      isColumnReorderEnabled: true,
    };
  }
  public async componentWillMount() {
    await this.getData();
  }

  public async getData() {
    const data: IDocument[] = [];
    let web = Web(this.props.webURL);
    const items: any[] = await web.lists
      .getByTitle("FluentUIDetailsListWithCustomItemColumn")
      .items.select("*", "CustomerName/EMail", "CustomerName/Title")
      .expand("CustomerName/ID")
      .get();
    console.log(items);
    await items.forEach(async (item) => {
      await data.push({
        CustomerName: item.CustomerName.Title,
        CustomerEmail: item.CustomerName.EMail,
        ProductName: item.ProductName,
        OrderDate: FormatDate(item.OrderDate),
        ProductDescription: item.ProductDescription,
        ProductImage:
          window.location.origin +
          item.ProductImage.match('"serverRelativeUrl":(.*),"id"')[1].replace(
            /['"]+/g,
            ""
          ),
      });
    });
    console.log(data);
    await this.setState({ Items: data });

    console.log(this.state.Items);
  }
  public _onRenderItemColumn = (
    item: IDocument,
    index: number,
    column: IColumn
  ): JSX.Element | string => {
    const src = item.ProductImage;

    switch (column.key) {
      case "ProductImage":
        return (
          <a href={item.ProductImage} target="_blank">
            <Image src={src} width={50} height={50} imageFit={ImageFit.cover} />
          </a>
        );
      case "ProductName":
        return (
          <span data-selection-disabled={true} style={{ whiteSpace: "normal" }}>
            {item.ProductName}
          </span>
        );

      case "CustomerName":
        return (
          <Link style={{ whiteSpace: "normal" }} href="#">
            {item.CustomerName}
          </Link>
        );

      case "CustomerEmail":
        return (
          <span style={{ whiteSpace: "normal" }}>{item.CustomerEmail}</span>
        );

      case "OrderDate":
        return (
          <span
            data-selection-disabled={true}
            className={mergeStyles({ height: "100%", display: "block" })}
          >
            {item.OrderDate}
          </span>
        );
      case "ProductDescription":
        return (
          <span data-selection-disabled={true} style={{ whiteSpace: "normal" }}>
            {item.ProductDescription}
          </span>
        );
      default:
        return <span>{item.CustomerName}</span>;
    }
  };
  public render(): React.ReactElement<IFluentUiDetailsListWithCustomItemColumnProps> {
    return (
      <div>
        <h1>Fluent UI DetailsList with Custom Item Column</h1>
        <DetailsList
          items={this.state.Items}
          columns={this.state.columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          onRenderItemColumn={this._onRenderItemColumn}
          selectionMode={SelectionMode.none}
        />
        <h3>Note : Image is clickable.</h3>
      </div>
    );
  }
}
export const FormatDate = (date): string => {
  var date1 = new Date(date);
  var year = date1.getFullYear();
  var month = (1 + date1.getMonth()).toString();
  month = month.length > 1 ? month : "0" + month;
  var day = date1.getDate().toString();
  day = day.length > 1 ? day : "0" + day;
  return month + "/" + day + "/" + year;
};
