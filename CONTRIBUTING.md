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

* Install [nodejs](https://nodejs.org) if you don't already have it.
* We recommend using [VS Code](https://code.visualstudio.com/) as your editor
* We also recommend installing the recommended extensions (recommended by VS Code once the folder is opened)

Clone the repo and run the following in the terminal/command line/powershell:

```bash
npm install
npm start
```

## Testing the preview NPM package
If you want to install the preview package just once, run this command:
```bash
npm install microsoft-graph-toolkit --registry https://pkgs.dev.azure.com/microsoft-graph-toolkit/_packaging/MGT/npm/registry/
```
If you want to be able to do it again later, by running ```npm install```, add the preview feed to your project's '.npmrc' file (create a '.npmrc' file if you don't have one already):
```
registry=https://pkgs.dev.azure.com/microsoft-graph-toolkit/_packaging/MGT/npm/registry/
always-auth=true
```
Then just install the package:
```bash
npm install microsoft-graph-toolkit
```

[TODO: Make sure everybody can access this feed publicly.]

## Contributing

### Finding issues you can help with
Looking for something to work on?
[Issues marked *good first issue*](https://github.com/microsoftgraph/microsoft-graph-toolkit/labels/good%20first%20issue)
are a good place to start.

You can also check [the *help wanted* tag](https://github.com/microsoftgraph/microsoft-graph-toolkit/labels/help%20wanted)
to find other issues to help with.

### Git workflow
The Microsoft Graph Toolkit uses the [GitHub flow](https://guides.github.com/introduction/flow/) where most development happens directly on the `master` branch. The `master` branch should always be in a healthy state which is ready for release.

If your change is complex, please clean up the branch history before submitting a pull request.
You can use [git rebase](https://docs.microsoft.com/en-us/azure/devops/repos/git/rebase#squash-local-commits) to group your changes into a small number of commits which we can review one at a time.

When completing a pull request, we will generally squash your changes into a single commit. Please let us know if your pull request needs to be merged as separate commits.

### Submitting a pull request and participating in code review
Writing a good description for your pull request is crucial to help reviewers and future maintainers understand your change. Make sure to complete the pull request template to avoid delays. More detail is better.
- [Link the issue you're addressing in the pull request](https://github.com/blog/957-introducing-issue-mentions). Each pull request must be linked to an issue.
- Describe *why* the change is being made and *why* you've chosen a particular solution.
- Describe any manual testing you performed to validate your change.
- Ensure the appropriate tests and documentation have been added

Please submit one pull request per issue. Large pull requests which have unrelated changes can be difficult to review.

After submitting a pull request, core members of the project will review your code. Each pull request must be validated by at least two core members before being merged

Often, multiple iterations will be needed to responding to feedback from reviewers.

### Quality assurance for new features
We encourage developers to follow the following guidance when submitting new features:
1. Ensure the appropriate tests have been added in the `src\test` folder. Run the tests and make sure they all pass.
1. Ensure the code is properly documented following the [tsdoc](https://github.com/Microsoft/tsdoc) syntax
1. Update the [documentation] when necessary
1. Follow the [accessibility guidance](https://developer.mozilla.org/en-US/docs/Web/Accessibility) for web development

### Testing
Your changes should include tests to verify new functionality wherever possible.

[[TODO - how to add and run tests]]

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
