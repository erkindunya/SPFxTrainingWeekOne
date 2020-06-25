$(document).ready(function () {
    $('#requests').DataTable({
        ajax: {
            url:
                'https://services.odata.org/V3/Northwind/Northwind.svc/Products?$format=json',
            headers: { Accept: 'application/json;odata=nometadata' },
            dataSrc: function (data) {
                return data.value.map(function (item) {
                    return [
                        item.ProductID,
                        item.ProductName,
                        item.QuantityPerUnit,
                        item.UnitPrice,
                        item.UnitsInStock,
                        //item.AssignedTo.Title
                    ];
                });
            },
        },
        columnDefs: [
            {
                targets: 4,
                //render: $.fn.dataTable.render.moment('YYYY/MM/DD')
            },
        ],
    });
});
