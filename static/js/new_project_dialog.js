class NewPrjectDialog {
    #dialog;
    #onOKProceed;
    #onOkHandler;
    #onCancelHandler;
    #onKeyDownHandler;

    constructor(projectsPath, onOk) {
        const dialog = $("#new-project-dialog");
        this.#dialog = dialog[0];
        this.#onOKProceed = onOk;

        this.#onOkHandler = this.#onOK.bind(this);
        this.#onCancelHandler = this.#onCancel.bind(this);
        this.#onKeyDownHandler = this.#onKeyDown.bind(this);

        dialog.on("click", this.#onClick.bind(this));
        $("#create-new-project-button").on("click", this.#onOkHandler);
        $("#create-new-project-cancel").on("click", this.#onCancelHandler);

        $("#projects-folder")
            .val(projectsPath)
            .on("keydown", this.#onKeyDownHandler);
    }

    show() {
        this.#dialog.showModal();
    }

    #onClick(e) {
        e.stopPropagation();
    }

    async #onOK() {
        $("#create-new-project-button").off("click", this.#onOkHandler);
        $("#create-new-project-cancel").off("click", this.#onCancelHandler);
        $("#projects-folder").off("keydown", this.#onKeyDownHandler);

        this.#dialog.close();

        const projectsPath = $("#projects-folder").val();

        if (projectsPath.includes("sandi")) {
            if (this.#onOKProceed)
                this.#onOKProceed(projectsPath);
        }
        else {
            alert("Incorrect projects path.")
        }
    }

    #onCancel() {
        $("#create-new-project-button").off("click", this.#onOkHandler);
        $("#create-new-project-cancel").off("click", this.#onCancelHandler);

        this.#dialog.close();
    }

    #onKeyDown(e) {
        if (e.originalEvent.key === "Enter")
            setTimeout(() => {this.#dialog.close(); this.#onOK();}, 0);
    }
}