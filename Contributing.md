# Contribute to this documentation

Thank you for your interest in our documentation! The best ways to help make this documentation better for everyone are to:

* [Report issues or request additional samples.](#report-issues)
* [Add or edit new samples.](#add-or-edit-samples)

## Report issues

Most of this repo's content is generated from internal files that are copied to **/generate-docs/script-inputs/excel.d.ts**. As such, documentation corrections are likely to be overwritten. Please report issues or make requests for additional documentation through [GitHub Issues](https://github.com/OfficeDev/office-scripts-docs-reference/issues). To do this, go to the **Feedback** section at the bottom of the affected article, then select **This page** to create a GitHub issue. Alternatively, create a new issue directly on [GitHub](https://github.com/OfficeDev/office-scripts-docs-reference/issues/new).

## Add or edit new samples

Samples are a critical tool for learning the Office Scripts API. We welcome sample scripts from the community to help demonstrate useful scripting scenarios.

### Adding a sample script to the GitHub repository

All the example code in this repository comes from one of the base sample **.yaml** files. The Office Scripts for Excel samples are in **[/docs/sample-scripts/excel-scripts.yaml](/docs/sample-scripts/excel-scripts.yaml)**. Use one of the methods described earlier in this guide to edit the file. You can also create an issue with the desired sample code and we'll add it to the documentation.

Samples are tied to particular APIs. Pick the API that your sample showcases and use it as the YAML key. A YAML key for the Excel sample file has this format: `'ExcelScript.'`*\<class-name\>*`#`*\<method-name\>*`:member(`*\<overload-number\>*`)`. The *\<overload-number\>* indicates which method is being referenced when there are multiple options. The value is usually `1`.

After the YAML key, add `:`, then add `  - |-` on a new line. This ensures the spacing in the example stays as is.

The following example shows a sample script for `Range.getValue`.

```YAML
'ExcelScript.Range#getValue:member(1)':
    - |-
    /**
     * This sample reads the value of A1 and prints it to the console.
     */
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();
    
        // Get the value of cell A1.
        let range = selectedSheet.getRange("A1");
        
        // Print the value of A1.
        console.log(range.getValue());
    }
```

Please add samples to the **.yaml** file in alphabetical order, based on the class and method names. This makes them easier to maintain.

### What makes a good sample?

A good sample has the following characteristics.

1. Descriptive - The sample has an introductory comment explaining the behavior and comments in the code as needed.
1. Useful - Some or all of the sample could be copied into an actual script and quickly modified to suit someone's needs.
1. Readable - The sample uses simple TypeScript, effective whitespace, and meaningful variable names.
1. Targeted - The sample showcases a specific API or scenario and only includes enough other code to be functional.

### When do samples get added to the actual documentation pages?

After the **sample-scripts** files are updated, a member of the team will run the documentation tooling and regenerate the API pages. Your new sample is added then.

### Contribute using GitHub

Use GitHub to contribute to this documentation without having to clone the repo to your desktop. This is the easiest way to create a pull request in this repository. Use this method to make a minor change that doesn't involve code changes.

Using this method allows you to contribute to one article at a time.

#### To contribute using GitHub

1. Find the article you want to contribute to on GitHub.
1. Once you are on the article in GitHub, sign in to GitHub (to get a free account, [Join GitHub](https://github.com/join)).
1. Choose the **pencil icon** (edit the file in your fork of this project) and make your changes in the **<>Edit file** window.
1. Scroll to the bottom and enter a description.
1. Choose **Propose file change** > **Create pull request**.

You now have successfully submitted a pull request. Pull requests are typically reviewed within 10 business days.

### Contribute using Git

If you're planning on making large changes, including updates to the documentation tooling, you should use Git to create a PR.

#### To contribute using Git

1. If you don't have a GitHub account, set one up at [GitHub](https://github.com/join).
1. After you have an account, install Git on your computer. Follow the steps in the [Set up Git] tutorial.
1. To submit a pull request using Git, follow the steps in [Use GitHub, Git, and this repository](#use-github-git-and-this-repository).
1. You will be asked to sign the Contributor's License Agreement if you are:

    * A member of the Microsoft Open Technologies group.
    * A contributor who doesn't work for Microsoft.

As a community member, you must sign the Contribution License Agreement (CLA) before you can contribute large submissions to a project. You only need to complete and submit the documentation once. Carefully review the document. You may be required to have your employer sign the document.

Signing the CLA does not grant you rights to commit to the main repository, but it does mean that the Office Developer and Office Developer Content Publishing teams will be able to review and approve your contributions. You are credited for your submissions.

## Use GitHub, Git, and this repository

**Note:** Most of the information in this section can be found in [GitHub Help] articles. If you're familiar with Git and GitHub, skip to the [Contribute and edit content](#contribute-and-edit-content) section for the specifics of the code/content flow of this repository.

### To set up your fork of the repository

1. Set up a GitHub account so you can contribute to this project. If you haven't done this, go to [GitHub](https://github.com/join) and do it now.
1. Install Git on your computer. Follow the steps in the [Set up Git] tutorial.
1. Create your own fork of this repository. To do this, at the top of the page,  choose the **Fork** button.
1. Copy your fork to your computer. To do this, open Git Bash. At the command prompt enter:

        git clone https://github.com/<your user name>/office-scripts-docs-reference.git

    Next, create a reference to the root repository by entering these commands:

        cd office-scripts-docs-reference
        git remote add upstream https://github.com/OfficeDev/office-scripts-docs-reference.git
        git fetch upstream

Congratulations! You've now set up your repository. You won't need to repeat these steps again.

### Contribute and edit content

To make the contribution process as seamless as possible, follow these steps.

#### To contribute and edit content

1. Create a new branch.
1. Add new content or edit existing content.
1. Submit a pull request to the main repository.
1. Delete the branch.

**Important**: Limit each branch to a single concept/article to streamline the work flow and reduce the chance of merge conflicts. Content appropriate for a new branch includes:

* A new article.
* Spelling and grammar edits.
* Applying a single formatting change across a large set of articles (for example, applying a new copyright footer).

#### To create a new branch

1. Open Git Bash.
1. At the Git Bash command prompt, type `git pull upstream main:<new branch name>`. This creates a new branch locally that is copied from the latest OfficeDev main branch.
1. At the Git Bash command prompt, type `git push origin <new branch name>`. This alerts GitHub to the new branch. You should now see the new branch in your fork of the repository on GitHub.
1. At the Git Bash command prompt, type `git checkout <new branch name>` to switch to your new branch.

#### Add new content or edit existing content

You navigate to the repository on your computer by using File Explorer. The repository files are in `C:\Users\<yourusername>\office-scripts-docs-reference`.

To edit files, open them in an editor of your choice and modify them. To create a new file, use the editor of your choice and save the new file in the appropriate location in your local copy of the repository. While working, save your work frequently.

The files in `C:\Users\<yourusername>\office-scripts-docs-reference` are a working copy of the new branch that you created in your local repository. Changing anything in this folder doesn't affect the local repository until you commit a change. To commit a change to the local repository, type the following commands in GitBash.

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

The `add` command adds your changes to a staging area in preparation for committing them to the repository. The period after the `add` command specifies that you want to stage all of the files that you added or modified, checking subfolders recursively. (If you don't want to commit all of the changes, you can add specific files. You can also undo a commit. For help, type `git add -help` or `git status`.)

The `commit` command applies the staged changes to the repository. The switch `-m` means you are providing the commit comment in the command line. The -v and -a switches can be omitted. The -v switch is for verbose output from the command, and -a does what you already did with the add command.

You can commit multiple times while you are doing your work, or you can commit once when you're done.

#### Submit a pull request to the main repository

When you're finished with your work and are ready to have it merged into the main repository, follow these steps.

#### To submit a pull request to the main repository

1. In the Git Bash command prompt, type `git push origin <new branch name>`. In your local repository, `origin` refers to your GitHub repository that you cloned the local repository from. This command pushes the current state of your new branch, including all commits made in the previous steps, to your GitHub fork.
1. On the GitHub site, navigate in your fork to the new branch.
1. Choose the **Pull Request** button at the top of the page.
1. Verify the Base branch is `OfficeDev/office-scripts-docs-reference@main` and the Head branch is `<your username>/office-scripts-docs-reference@<branch name>`.
1. Choose the **Update Commit Range** button.
1. Add a title to your pull request, and describe all the changes you're making.
1. Submit the pull request.

One of the site administrators will process your pull request. Your pull request will surface on the [OfficeDev/office-scripts-docs-reference](https://github.com/OfficeDev/office-scripts-docs-reference/pulls) site under **Pull requests**. When the pull request is accepted, the issue will be resolved.

#### Create a new branch after merge

After a branch is successfully merged (that is, your pull request is accepted), don't continue working in that local branch. This can lead to merge conflicts if you submit another pull request. To do another update, create a new local branch from the successfully merged upstream branch, and then delete your initial local branch.

For example, if your local branch X was successfully merged into the OfficeDev/office-scripts-docs-reference main branch and you want to make additional updates to the content that was merged. Create a new local branch, X2, from the OfficeDev/office-scripts-docs-reference main branch. To do this, open GitBash and execute the following commands.

    cd office-scripts-docs-reference
    git pull upstream main:X2
    git push origin X2

You now have local copies (in a new local branch) of the work that you submitted in branch X. The X2 branch also contains all the work other writers have merged, so if your work depends on others' work (for example, shared images), it is available in the new branch. You can verify that your previous work (and others' work) is in the branch by checking out the new branch...

    git checkout X2

...and verifying the content. (The `checkout` command updates the files in `C:\Users\<yourusername>\office-scripts-docs-reference` to the current state of the X2 branch.) Once you check out the new branch, you can make updates to the content and commit them as usual. However, to avoid working in the merged branch (X) by mistake, it's best to delete it (see the following **Delete a branch** section).

#### Delete a branch

Once your changes are successfully merged into the main repository, delete the branch you used because you no longer need it.  Any additional work should be done in a new branch.  

#### To delete a branch

1. In the Git Bash command prompt, type `git checkout main`. This ensures that you aren't in the branch to be deleted (which isn't allowed).
1. Next, at the command prompt, type `git branch -d <branch name>`. This deletes the branch on your computer only if it has been successfully merged to the upstream repository. (You can override this behavior with the `–D` flag, but first be sure you want to do this.)
1. Finally, type `git push origin :<branch name>` at the command prompt (a space before the colon and no space after it).  This will delete the branch on your github fork.  

Congratulations, you have successfully contributed to the project!

## FAQ

### How do I get a GitHub account?

Fill out the form at [Join GitHub](https://github.com/join) to open a free GitHub account.

### Where do I get a Contributor's License Agreement?

You will automatically be sent a notice that you need to sign the Contributor's License Agreement (CLA) if your pull request requires one.

As a community member, **you must sign the Contribution License Agreement (CLA) before you can contribute large submissions to this project**. You only need complete and submit the documentation once. Carefully review the document. You may be required to have your employer sign the document.

### What happens with my contributions?

When you submit your changes, via a pull request, our team will be notified and will review your pull request. You will receive notifications about your pull request from GitHub; you may also be notified by someone from our team if we need more information. If your pull request is approved, we'll update the documentation. We reserve the right to edit your submission for legal, style, clarity, or other issues.

### Can I become an approver for this repository's GitHub pull requests?

Currently, we are not allowing external contributors to approve pull requests in this repository.

### How soon will I get a response about my change request?

Pull requests are typically reviewed within 10 business days.

## More resources

* To learn more about using Git and GitHub, first check out the [GitHub Help].

[GitHub Help]: https://docs.github.com
[Set up Git]: https://docs.github.com/get-started/quickstart/set-up-git
