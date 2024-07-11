class Logger {
    constructor() {
        this.loggerObject = null;
    }

    setLoggerInstance(loggerInstance) {
        this.loggerObject = loggerInstance;
    }

    debug(params) {
        if (this.loggerObject) {
            this.loggerObject.debug(params);
        }
    }

    info(params) {
        if (this.loggerObject) {
            this.loggerObject.info(params);
        }
    }
}

module.exports = {
    Logger
};