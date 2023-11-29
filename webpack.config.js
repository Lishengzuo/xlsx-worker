const path = require("path");

module.exports = {
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
};
