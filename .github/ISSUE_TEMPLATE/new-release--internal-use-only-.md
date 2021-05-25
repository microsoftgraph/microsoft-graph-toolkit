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

#### Release
  - [ ] Ensure `main` is up to date with `release/latest`
  - [ ] Ensure `./package.json` version in `main` is correct
  - [ ] Merge `main` in `release/latest` - this will invoke the release workflow
  - [ ] Approve workflow to publish npm packages
  - [ ] Validate npm packages are published
  - [ ] Validate mgt.dev is updated
  - [ ] Validate GitHub release contains `mgt-spfx` asset
  - [ ] Increment `./package.json` version in main for building preview packages 
