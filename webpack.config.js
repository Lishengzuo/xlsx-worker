const path = require("path");

module.exports = {
  entry: "./src/index.js",
  output: {
    filename: "bundle.js",
    path: path.resolve(__dirname, "dist"),
  },
  resolve: {
    alias: {
      "worker-loader": path.resolve(__dirname, "node_modules/worker-loader"),
    },
  },
  module: {
    rules: [
      {
        test: /\.worker\.js$/,
        use: { loader: "worker-loader" },
      },
    ],
  },
  experiments: {
    topLevelAwait: false,
  },
  output: {
    globalObject: "self",
  },
};
