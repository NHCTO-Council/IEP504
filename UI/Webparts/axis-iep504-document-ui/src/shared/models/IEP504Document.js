var IEP504Document = (function () {
    function IEP504Document() {
    }
    IEP504Document.prototype.SetHeaderRow = function () {
        try {
            this.headerRow =
                this.student.lastName +
                    ", " +
                    this.student.firstName +
                    " (ID: " +
                    this.student.id +
                    ")";
        }
        catch (err) {
            return "Error: " + err;
        }
    };
    return IEP504Document;
}());
export { IEP504Document };
//# sourceMappingURL=IEP504Document.js.map