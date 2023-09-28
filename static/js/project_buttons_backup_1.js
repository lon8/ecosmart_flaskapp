const EXCEL_ERROR = "Error processing Excel file.";

function alertNeedSeelectedProject() {
    alert("Please select a project.");
}

function createDisclaimerPDF(e) {
    e.preventDefault();

    const selectedRows = gridOptions.api.getSelectedRows();

    if (selectedRows.length) {
        window.open(`/create?identifier=${selectedRows[0][window.PATH_COLUMN]}`, '_blank').focus();
    }
    else
        alertNeedSeelectedProject();
}

function createNewProject(e) {
    e.preventDefault();
    const newProjectDialog = new NewPrjectDialog(window.projectsPath, newProjectsPath => {
        window.projectsPath = newProjectsPath;
        $("#newProjectFile").click();
    });
    newProjectDialog.show();
}

function processAuditReportStandalone(e) {
    e.preventDefault();
    $("#auditReportFileStandalone").click();
}

function createCustomProposal(e) {
    e.preventDefault();
    $("#proposalCustomFile").click();
}

function createPhotosCheckList(e) {
    e.preventDefault();
    $("#photosCheckListFile").click();
}

function createScopePDF(e) {
    e.preventDefault();
    $("#scopeFile").click();
}

async function onNewProjectCreated(response) {
    gridOptions.api.setRowData(JSON.parse(await response.data.text())); // blob
    hide_base_loading();
}

function createProjectRowHandler(url) {
    return async function(e) {
        e.preventDefault();

        const selectedRows = gridOptions.api.getSelectedRows();

        if (selectedRows.length) {
            show_base_loading();
            try {
                await axios.post(url, selectedRows);
            }
            catch (e) {
                await checkBackendError(e);
            }
            finally {
                hide_base_loading();
            }
        }
        else
            alertNeedSeelectedProject();
    }
}

function createExcelUploadHandler(url, postProcess) {
    return function (e) {
        postProcess = postProcess || (() => {});

        var reader = new FileReader();
        var excelFile = e.target.files[0];

        reader.readAsArrayBuffer(excelFile);

        reader.onload = async function () {
            show_base_loading();

            var formData = new FormData();
            var fileBlob = new Blob([reader.result], {type: "application/octet-stream"});

            formData.append("file_name", excelFile.name);
            formData.append("projects_folder", projectsPath);
            formData.append("content", fileBlob);

            try {
                await axios.post(url, formData, {responseType: "blob"}).then(postProcess);
            }
            catch (e) {
                checkBackendError(e);
            }
            finally {
                hide_base_loading();
            }
        };

        reader.onerror = function () {
            console.error(reader.error);
        };
    }
}

async function checkBackendError(e) {
    if (!e.response || e.response.status === 500) {

        if (e.response.data.display) {
            alert(e.response.data.message);
        }
        else {
            alert(EXCEL_ERROR);
            console.error(e);
            // const errorText = JSON.parse(await e.data.text());
            // console.error(errorText);
        }
    }
}

function processExcelDownload(response) {
    var fileName = response.headers["x-file-name"];
    const downloadUrl = URL.createObjectURL(response.data);
    const a = document.createElement("a");

    a.style.display = "none";
    a.href = downloadUrl;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    hide_base_loading();
    //URL.revokeObjectURL(downloadUrl);
}

function configureButtons() {
    $("#createPDFLink").on("click", createDisclaimerPDF);

    $("#newProject").on("click", createNewProject);
    $("#newProjectFile").on("change",
        createExcelUploadHandler("/api/create_new_project_from_raw_report", onNewProjectCreated));

    $("#reprocessAuditReport").on("click", createProjectRowHandler("/api/reprocess_audit_report"));

    $("#processAuditReportWithDefaults").on("click", createProjectRowHandler("/api/reprocess_audit_report_with_defaults"));

    $("#createProposalPDF").on("click", createProjectRowHandler("/api/create_proposal_pdf"));

    $("#createCustomProposal").on("click", createCustomProposal);
    $("#proposalCustomFile").on("change",
        createExcelUploadHandler("/api/create_proposal_custom", processExcelDownload));

    $("#processAuditReportStandalone").on("click", processAuditReportStandalone);
    $("#auditReportFileStandalone").on("change",
        createExcelUploadHandler("/api/process_audit_report_standalone", processExcelDownload));

    $("#createPhotosCheckList").on("click", createPhotosCheckList);
    $("#photosCheckListFile").on("change",
        createExcelUploadHandler("/api/create_photos_check_list", processExcelDownload));

    $("#createProjectScopePDF").on("click", createProjectRowHandler("/api/create_project_scope_pdf"));

    $("#createScopePDF").on("click", createScopePDF);
    $("#scopeFile").on("change",
        createExcelUploadHandler("/api/create_scope_pdf", processExcelDownload));

    $("#createCLIP").on("click", createProjectRowHandler("/api/create_clip"));

    $("#createMaterials").on("click", createProjectRowHandler("/api/create_materials"));
    $("#createMaterialsCombined").on("click", createProjectRowHandler("/api/create_materials_combined"));
}