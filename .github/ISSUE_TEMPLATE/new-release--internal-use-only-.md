---
name: New Release (internal use only)
about: This template is used for tracking releases
title: "[RELEASE] X.X"
labels: release-task
assignees: ''

---

This issue tracks the publishing progress for a release. It will get updated as tasks are completed

- [ ] Write blog post
- [ ] Prep the release branch
  - [ ] Merge main into 'release/latest'
  - [ ] Update the package versions to new version
- [ ] Publish release to GitHub
  - [ ] Write release notes
    - [ ] Write summary for top of release notes
  - [ ] Publish release notes
- [ ] Publish docs
  - [ ] Create docs PR
  - [ ] Merge docs PR
- [ ] Publish storybook from release branch
- [ ] Publish NPM package from release branch
- [ ] Merge release branch to main
- [ ] Publish release blog post
