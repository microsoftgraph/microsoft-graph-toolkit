# Contributing to the Microsoft Graph Toolkit

We strongly encourage community feedback and contributions.

## Our foundation

The foundation of the **Microsoft Graph Toolkit** is simplicity.

A developer should be able to quickly and easily learn to use the component/API.

Simplicity and a low barrier to entry are must-have features of every component and API. If you have any second thoughts about the complexity of a design, it is almost always much better to cut the feature from the current release and spend more time to get the design right for the next release.

You can always add to an API, you cannot ever remove anything from one. If the design does not feel right, and you ship it anyway, you are likely to regret having done so.

## Do you have a question?

For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

## Reporting issues and suggesting new features

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature request.

We highly recommend you browse existing issues first.

## Setting up the project

- Install [nodejs](https://nodejs.org) if you don't already have it.
- We recommend using [VS Code](https://code.visualstudio.com/) as your editor
- We also recommend installing the recommended extensions (recommended by VS Code once the folder is opened)

Clone the repo and run the following in the terminal/command line/powershell:

```bash
npm install
npm start
```

## Creating a new component (quick start)

The best way to get started with a new component is to use our snippets for scaffolding:

> Note: The steps below assume you are using Visual Studio Code and you've installed all the recommended extensions.

1. Chose a name for your component. The component must be prefixed with `mgt`. For example `mgt-component`

1. Create a new folder for your new component under `src` \ `components` and name it with the name of your component.

1. Create a new typescript file in your new folder with the same name. Ex: `mgt-component.ts`

1. Open the file and use the `mgtnew` snippet to scaffold the new component.

    > To use a snippet, start typing the name of the snippet (`mgtnew`) and press `tab`

1. Tab through the generated code to set the name of your component. 

1. Add your code!

<!-- ## Testing the preview NPM package

If you want to install the preview package just once, run this command:

```bash
npm install @microsoft/mgt --registry https://pkgs.dev.azure.com/microsoft-graph-toolkit/_packaging/MGT/npm/registry/
```

If you want to be able to do it again later, by running `npm install`, add the preview feed to your project's '.npmrc' file (create a '.npmrc' file if you don't have one already):

```
registry=https://pkgs.dev.azure.com/microsoft-graph-toolkit/_packaging/MGT/npm/registry/
always-auth=true
```

Then just install the package:

```bash
npm install @microsoft/mgt
```

[TODO: Make sure everybody can access this feed publicly.] -->

## Contributing

### Finding issues you can help with

Looking for something to work on?
[Issues marked _good first issue_](https://github.com/microsoftgraph/microsoft-graph-toolkit/labels/good%20first%20issue)
are a good place to start.

You can also check [the _help wanted_ tag](https://github.com/microsoftgraph/microsoft-graph-toolkit/labels/help%20wanted)
to find other issues to help with.

### Git workflow

The Microsoft Graph Toolkit uses the [GitHub flow](https://guides.github.com/introduction/flow/) where most development happens directly on the `master` branch. The `master` branch should always be in a healthy state which is ready for release.

If your change is complex, please clean up the branch history before submitting a pull request.
You can use [git rebase](https://docs.microsoft.com/en-us/azure/devops/repos/git/rebase#squash-local-commits) to group your changes into a small number of commits which we can review one at a time.

When completing a pull request, we will generally squash your changes into a single commit. Please let us know if your pull request needs to be merged as separate commits.

### Submitting a pull request and participating in code review

Writing a good description for your pull request is crucial to help reviewers and future maintainers understand your change. Make sure to complete the pull request template to avoid delays. More detail is better.

- [Link the issue you're addressing in the pull request](https://github.com/blog/957-introducing-issue-mentions). Each pull request must be linked to an issue.
- Describe _why_ the change is being made and _why_ you've chosen a particular solution.
- Describe any manual testing you performed to validate your change.
- Ensure the appropriate tests and [documentation](https://github.com/microsoftgraph/microsoft-graph-docs/tree/master/concepts/toolkit) have been added

Please submit one pull request per issue. Large pull requests which have unrelated changes can be difficult to review.

After submitting a pull request, core members of the project will review your code. Each pull request must be validated by at least two core members before being merged

Often, multiple iterations will be needed to responding to feedback from reviewers.

### Quality assurance for new features

We encourage developers to follow the following guidance when submitting new features:

1. Ensure the appropriate tests have been added in the `src\test` folder. Run the tests and make sure they all pass.
1. Ensure the code is properly documented following the [tsdoc](https://github.com/Microsoft/tsdoc) syntax
1. Update the [documentation](https://github.com/microsoftgraph/microsoft-graph-docs/tree/master/concepts/toolkit) when necessary
1. Follow the [accessibility guidance](https://developer.mozilla.org/en-US/docs/Web/Accessibility) for web development


### Accessibility Guidelines

New features and components should folow the following accessibility implementation guidelines:

(for ease of use)
1. Visit the following location: https://accessibilityinsights.io/en/
2. Install the extension, and test

**required**:
- [ ] `aria-label`- meaningful text should have identifiable labels for screen readers
- [ ] `label`- appropriate labels for custom elements (e.g. people-picker's label is input)
- [ ] `tab-index/focus`- components that are interactive or display information should be **visibilly** navigatable by `tab` key control. Additional information in the aria label should be displayed when this feature is used.
- [ ] `alt`- any `<img>` tag should contain `alt` text as well


<!-- ### Testing

Your changes should include tests to verify new functionality wherever possible.

[[TODO - how to add and run tests]] -->

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
