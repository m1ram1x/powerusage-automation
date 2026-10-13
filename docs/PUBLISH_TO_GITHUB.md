# Publish this project on GitHub

This guide prepares **source code** for a repository. The repository has not been created or uploaded automatically.

## 1. Use the clean folder

Extract `FMS_GitHub_Ready.zip` and open its `powerusage-automation` folder. This is the folder containing `README.md`, `main.py`, `Dockerfile` and `.gitignore`.

Upload its **contents**, not the ZIP itself and not an additional enclosing `github-ready` folder. The README should be at the repository root so GitHub displays it.

The package includes:

- Current Python source, including Telegram document delivery and date-based report filenames.
- Dockerfile and Linux/Windows launchers.
- Excel template in `templates`.
- Empty `input`/`output` folders retained through `.gitkeep`.
- Synthetic demo data and offline test fixtures.
- English usage, Telegram, troubleshooting and publishing guides.
- Git ignore rules, line-ending rules and a GitHub Actions test workflow.

Do not upload a real `telegram.env`, private CSV files, output Excel reports, virtual environments or built Docker archives. The original project archive, extracted Docker layers, history and credentials are not part of this clean source package.

## 2. Create a repository

Sign in to GitHub and choose **New repository**. A suggested name is `powerusage-automation`; you can choose another name. Suggested description:

```text
PDU CSV-to-Excel batch reporting with Docker, date-based reports and Telegram alerts/document delivery.
```

Choose Public or Private deliberately. For a Git-based first upload, create the repository **without** adding a README, .gitignore or license: these source files already exist locally. You may add a chosen license later.

[GitHub repository creation guide](https://docs.github.com/en/repositories/creating-and-managing-repositories/creating-a-new-repository).

## 3. Upload using Git (recommended)

Install Git if necessary. Open Terminal, Git Bash or PowerShell **inside the clean project folder**.

```bash
git init
git branch -M main
git add .
git status --short
```

Inspect the staged filenames. You should see source, documentation, template, demos and tests. There should be **no** actual `telegram.env`, private `input/*.csv`, `output/*.xlsx`, `.venv`, `.tar`, `.run` or old ZIP files. Git ignores protect a new upload; they do not remove a file that was previously tracked.

If Git requests an identity, configure your own name and GitHub-associated email locally for this repository:

```bash
git config user.name "YOUR_NAME"
git config user.email "YOUR_GITHUB_EMAIL"
```

Replace the example identity strings with your own values. Then:

```bash
git commit -m "Add Power Usage Automation"
git remote add origin YOUR_REPOSITORY_URL
git push -u origin main
```

Replace `YOUR_REPOSITORY_URL` with the HTTPS or SSH URL shown on your new repository's page. For HTTPS authentication, use Git's credential manager or GitHub's supported sign-in method. Do not put a password or access token directly in a command or remote URL.

## 4. Browser-upload alternative

Open your repository, select **Add file → Upload files** (or its initial upload link), and upload the clean folder's contents. Include hidden files/folders: `.gitignore`, `.dockerignore`, `.gitattributes` and `.github/workflows/ci.yml`. In Ubuntu's file manager, Ctrl+H shows hidden files.

**Browser uploads do not apply your local .gitignore.** Use only the supplied clean source folder, not your working folder containing secrets or output files. Verify the file list before committing the upload.

[GitHub file-upload guide](https://docs.github.com/en/repositories/working-with-files/managing-files/adding-a-file-to-a-repository).

## 5. Verify the published project

Open the GitHub repository and check that:

1. The English README appears on the main page and its relative documentation links work.
2. `.github/workflows/ci.yml` is present. Look at **Actions → Tests and Docker** for the test/build result.
3. `telegram.env.example` contains only blank values, and no real credentials or measurement files were uploaded.
4. **Code → Download ZIP** provides a source copy with the template and demo files.

No live Telegram credentials are required by CI. Its network requests download public dependencies/images only; Telegram requests in tests are mocked. The Docker smoke run processes synthetic CSV data with Telegram disabled.

## 6. Optional: publish a ready image as a Release

Keep generated binary installers and Docker images **out of Git history**. GitHub's browser repository uploads have a 25 MiB per-file limit, and normal Git blocks files over 100 MiB. Use **Releases** for distributable binaries. [GitHub large-file guidance](https://docs.github.com/en/repositories/working-with-files/managing-large-files/about-large-files-on-github).

Build the image from this source version on a machine with Docker:

```bash
docker build -t powerusage-automation:local .
docker save -o powerusage-image.tar powerusage-automation:local
```

On Linux, prepend `sudo` if needed, then make the exported file owned by your account before copying it. Create a GitHub Release, choose a tag such as `v1.0.0`, and attach the exported tar file. Include the target architecture in the release notes, such as Linux amd64/Intel/AMD, and state which source commit was used to build it. Prefer assets rebuilt from that tagged commit rather than mixing old images with newer source.

Recipients can load it with `docker load -i powerusage-image.tar`. The launchers recognize the `powerusage-automation:local` tag and use the loaded image; otherwise they build from source.

A real `telegram.env` must never be included in the image or release archive. The supplied Dockerfile/.dockerignore allowlist keeps local input, output and settings outside the image.

## 7. License and later updates

A GitHub repository can be created without a LICENSE file. Select a license yourself if you want to specify reuse terms; none was assigned automatically in this package. Follow [GitHub's licensing guide](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/licensing-a-repository).

For later source changes:

```bash
git add .
git status --short
git commit -m "Describe your change"
git push
```

Review staged files before each push. Rebuild Docker images after changing code/dependencies, and update release assets when you publish a new version.
