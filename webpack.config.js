const path = require("node:path")
const CopyWebpackPlugin = require("copy-webpack-plugin")
const devCerts = require("office-addin-dev-certs")

module.exports = async () => {
  const httpsOptions = await devCerts.getHttpsServerOptions()

  return {
    entry: {
      taskpane: path.resolve(__dirname, "src/taskpane/scripts/taskpane.js"),
      "auth-start": path.resolve(__dirname, "src/taskpane/scripts/auth-start.js"),
      "auth-callback": path.resolve(__dirname, "src/taskpane/scripts/auth-callback.js"),
    },
    output: {
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          { from: "src/taskpane/pages/taskpane.html", to: "taskpane.html" },
          { from: "src/taskpane/pages/auth-start.html", to: "auth-start.html" },
          { from: "src/taskpane/pages/auth-callback.html", to: "auth-callback.html" },
          { from: "src/taskpane/styles/taskpane.css", to: "taskpane.css" },
          { from: "assets", to: "assets" },
          { from: "manifest.xml", to: "manifest.xml" },
        ],
      }),
    ],
    devtool: "source-map",
    devServer: {
      hot: false,
      port: 3000,
      server: {
        type: "https",
        options: httpsOptions,
      },
      static: {
        directory: path.resolve(__dirname, "dist"),
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
  }
}
