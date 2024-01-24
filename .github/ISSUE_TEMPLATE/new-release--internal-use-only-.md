---
name: New Release (internal use only)
about: This template is used for tracking releases
title: "[RELEASE] X.X"
labels: release-task
assignees: ''

---

This issue tracks the publishing progress for a release. It will get updated as tasks are completed

#### Blog Post
  - [ ] Author and stage
  - [ ] Publish

#### GitHub Release Notes
  - [ ] Author draft release notes
  - [ ] Author release notes summary
  - [ ] Publish

#### Docs
  - [ ] Create docs PR
  - [ ] Get docs PR reviewed and approved
  - [ ] Merge docs PR

#### Design
  - [ ] Design review

#### Bug Bash
  - [ ] Bash bugs
  - [ ] mgt.dev is embedable

#### Release
  - [ ] Ensure `main` is up to date with `release/latest`
  - [ ] Ensure `./package.json` version in `main` is correct
  - [ ] Merge `main` in `release/latest` - this will invoke the release workflow
  - [ ] Approve workflow to publish npm packages
  - [ ] Validate npm packages are published
  - [ ] Validate mgt.dev is updated
  - [ ] Validate https://aka.ms/ge and https://aka.ms/mgt/docs are embedding the latest documentation
  - [ ] Publish GitHub release
  - [ ] Close the current milestone
  - [ ] Increment `./package.json` version in main for building preview packages 
