module.exports = function handler(_request, response) {
    response
        .status(200)
        .setHeader("Content-Type", "text/plain; charset=utf-8");
    response.end("Deployment Form Generator service is running");
};
