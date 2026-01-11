<p align="center">
    <!-- .NET -->
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://avatars.githubusercontent.com/u/9141961?s=200&v=4" alt=".NET Logo">
      <img src="https://avatars.githubusercontent.com/u/9141961?s=200&v=4" width="225" alt=".NET Logo">
    </picture>
    <h1 align="center">msguru</h1>
</p>

[![Renovate](https://img.shields.io/badge/Renovate-enabled-brightgreen?logo=renovate&logoColor=1A1F6C)][renovate]
[![PreCommit](https://img.shields.io/badge/PreCommit-enabled-brightgreen?logo=precommit&logoColor=FAB040)][precommit]

A collection of open-source **[MIT-licensed][license]** command-line tools built with **.NET 10**, designed for inspecting, converting, extracting, and automating operations on **Microsoft product files** — including **Outlook MSG/PST**, **Word DOCX**, and additional formats built on Office technologies.
`msguru` combines the power of **System.CommandLine** with **Microsoft Office interop libraries** to provide a unified, scriptable, and extensible interface for working with complex enterprise file formats.

Whether you need to parse email metadata, batch-export attachments, analyze PST archives, or transform document contents, `msguru` aims to streamline your workflow with a clean and modular CLI experience.

## ✨ TL;DR

```shell
# restore .NET tools & dependencies
dotnet restore

# view available commands and options
dotnet run  --project src/Main -- --help

```

### 🔃 Contributing

Refer to our [documentation for contributors][contributing] for contributing guidelines, commit message
formats and versioning tips.

### 📥 Maintainers

This project is owned and maintained by [Ad Noctem Collective][github] refer to the [`AUTHORS`][authors] or
[`CODEOWNERS`][codeowners] for more information. You may also use the linked contact details to reach out directly.

---

### 📜 License

**[MIT][license]**

### ©️ Copyright

Assets provided by [Microsoft &reg;][microsoft].

<!-- INTERNAL REFERENCES -->

<!-- File references -->

[license]: LICENSE
[authors]: .github/AUTHORS
[codeowners]: .github/CODEOWNERS
[contributing]: docs/CONTRIBUTING.md

<!-- General links -->

[github]: https://github.com/adnoctem
[microsoft]: https://microsoft.com

<!-- [typescript]: https://www.typescriptlang.org/ -->

<!-- Third-party -->

[renovate]: https://renovatebot.com/
[precommit]: https://pre-commit.com/
