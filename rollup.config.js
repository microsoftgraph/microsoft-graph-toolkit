import postcss from "rollup-plugin-postcss";
import resolve from "rollup-plugin-node-resolve";
import commonJS from "rollup-plugin-commonjs";

const core_opts = {
  external: ["@microsoft/teams-js"],
  plugins: [
    postcss({
      inject: false,
      use: ["sass"]
    }),
    resolve(),
    commonJS(),
  ]
};

export default [
  {
    ...core_opts,
    input: "src/index.js",
    output: {
      dir: "bin/full-bundle/",
      format: "esm"
    }
  }
  // {
  //   ...core_opts,
  //   input: "src/components/mgt-agenda/mgt-agenda",
  //   output: {
  //     file: "bin/components/mgt-agenda.js",
  //     format: "esm"
  //   }
  // },
  // {
  //   ...core_opts,
  //   input: "src/components/mgt-login/mgt-login.js",
  //   output: {
  //     file: "bin/components/mgt-login.js",
  //     format: "esm"
  //   }
  // },
  // {
  //   ...core_opts,
  //   input: "src/components/mgt-person/mgt-person.js",
  //   output: {
  //     file: "bin/components/mgt-person.js",
  //     format: "esm"
  //   }
  // }
];
