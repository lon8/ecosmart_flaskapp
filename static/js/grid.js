function dataChanged() {
    document.getElementById('dataChanged').style.display = "block";
    document.getElementById('reset_link').style.display = "none";
    document.getElementById('show_hide_link').style.display = "none";
}

function dataPristine() {
    document.getElementById('dataChanged').style.display = "none";
    document.getElementById('reset_link').style.display = "inline";
    document.getElementById('show_hide_link').style.display = "inline";
}

function doSaveColumnState() {
    if (!window.ready_to_save) return;
    console.log('save');
    var data = gridOptions.columnApi.getColumnState();
    window.localStorage.setItem('fl1616099_cache', JSON.stringify(data));
}

var saveColumnState = _.debounce(doSaveColumnState);

gridOptions = {
    animateRows: true, //pagination: true,
    rowSelection: 'multiple',

    defaultColDef: {
        editable: true, sortable: true, filter: true, resizable: true,
    },

    onGridReady: function (params) {
    },

    onFirstDataRendered: function () {
        var columnState = null;
        try {
            columnState = JSON.parse(window.localStorage.getItem('fl1616099_cache'));
        } catch (e) {
            console.error(e);
        }
        if (columnState) {
            gridOptions.columnApi.applyColumnState({
                state: columnState, applyOrder: true
            });
        }
        window.ready_to_save = true;
    },

    onCellValueChanged: function () {
        dataChanged();
    },

    onFilterChanged: saveColumnState,
    onSortChanged: saveColumnState,
    onColumnEverythingChanged: saveColumnState,
    onColumnVisible: saveColumnState,
    onColumnPinned: saveColumnState,
    onColumnResized: saveColumnState,
    onColumnMoved: saveColumnState,
}

getCurrentData = function () {
    var d = [];
    gridOptions.api.forEachNode(function (r) {
        d.push(r.data);
    })
    return d;
}

var postData = function () {
    var data = getCurrentData();
    show_base_loading();
    axios.post('/api/data', {data: data}).then(function (resp) {
        if (resp.data.status === 'ok') {
        }
        location.reload();
    }).catch(function () {
        hide_base_loading();
        alert('Error occured');
    });
}

var table;

function clearColumnCache() {
    if (confirm("Do you really want to clear column cache?")) {
        window.localStorage.setItem('fl1616099_cache', null);
        location.reload();
    }
}

function onCheckboxChange(colId) {
    var el = document.getElementById("checkbox_" + colId);
    window.columnDialogData[colId] = el.checked;
}

function editHiddenColumns() {
    var data = gridOptions.columnApi.getColumnState();
    var html = '<div style="text-align: left">';
    window.columnDialogData = {};
    for (var i = 0; i < data.length; i++) {
        var item = data[i];
        var column = gridOptions.columnApi.getColumn(item.colId);
        window.columnDialogData[item.colId] = !item.hide;
        html += '<label style="display: block; margin-bottom: 10px"><input id="checkbox_' + item.colId + '" type="checkbox" ' + (item.hide? '': 'checked="checked"') + ' onchange="onCheckboxChange(\'' + item.colId + '\')">' + item.colId + '</label>'
    }
    html += '</div>';
    Swal.fire({
        title: null, inputPlaceholder: 'Select visible columns', showCancelButton: true, html: html,
    }).then(function (result) {
        if (result.value) {
            for (var i = 0; i < data.length; i++) {
                var item = data[i];
                if (window.columnDialogData[item.colId] !== !item.hide) {
                    gridOptions.columnApi.setColumnVisible(item.colId, window.columnDialogData[item.colId]);
                }
            }
        }
    });
}

document.addEventListener("DOMContentLoaded", function () {
    var gridDiv = document.querySelector('#grid');
    axios.get("/api/data").then(function (resp) {
        var data = resp.data;

        //document.getElementById('title').textContent = data.sheet_name;
        var columnDefs = data.columns.map(function (col) {
            return {field: col};
        });
        gridOptions.columnDefs = columnDefs;
        gridOptions.rowData = data.data;
        window.table = new agGrid.Grid(gridDiv, gridOptions);
        document.getElementById("add_row_link").addEventListener('click', function () {
            gridOptions.api.applyTransaction({
                add: [{}],
            });
            dataChanged();
        });
        document.getElementById("save_link").addEventListener('click', postData);
        document.getElementById("reset_link").addEventListener('click', clearColumnCache);
        document.getElementById("show_hide_link").addEventListener('click', editHiddenColumns);

        configureButtons();
    });
});

function show_base_loading() {
    document.getElementById("base-loading-shim").style.display = "flex";
}

function hide_base_loading() {
    document.getElementById("base-loading-shim").style.display = "none";
}
