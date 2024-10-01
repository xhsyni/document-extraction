document.addEventListener('DOMContentLoaded', function () {
    const folderInput = document.getElementById('pdfFolder');
    const fileList = document.getElementById('file-list');

    folderInput.addEventListener('change', function () {
        fileList.innerHTML = '';

        for (const file of folderInput.files) {
            const listItem = document.createElement('div');
            listItem.classList.add('file-item');

            const fileName = document.createElement('span');
            fileName.textContent = file.name;

            const deleteButton = document.createElement('button')
            deleteButton.textContent = 'Delete';
            deleteButton.type = 'button';

            deleteButton.addEventListener('click', function () {
                removeFile(file);
                fileList.removeChild(listItem)
            })
            listItem.appendChild(deleteButton);
            listItem.appendChild(fileName);
            fileList.appendChild(listItem);
        }
    });

    const fileInput = document.getElementById('pdfFile');
    const filePath = document.getElementById('file-path');
    fileInput.addEventListener('change', function () {
        filePath.innerHTML = '';

        filePath.textContent = fileInput.files[0].name
    })

    function removeFile(fileToRemove) {
        const dataTransfer = new DataTransfer();
        for (const file of folderInput.files) {
            if (file !== fileToRemove) {
                dataTransfer.items.add(file);
            }
        }
        folderInput.files = dataTransfer.files;
    }

    // Handle the file and folder button
    const fileContainer = document.querySelector('.mainform-filecontainer');
    const folderContainer = document.querySelector('.mainform-foldercontainer');

    const uploadFileButton = document.getElementById("file-upload")
    const uploadFolderButton = document.getElementById("folder-upload")

    uploadFileButton.classList.add('show');
    uploadFileButton.addEventListener('click', function () {
        fileContainer.style.display = 'block';
        folderContainer.style.display = 'none';
        uploadFileButton.classList.add('show');
        uploadFolderButton.classList.remove('show');
    })

    uploadFolderButton.addEventListener('click', function () {
        folderContainer.style.display = 'block';
        fileContainer.style.display = 'none';
        uploadFileButton.classList.remove('show');
        uploadFolderButton.classList.add('show');
    })

    const fileForm = document.querySelector('.mainform-filecontainer form');
    const folderForm = document.querySelector('.mainform-foldercontainer form');

    const submitFileButton = document.getElementById('submit-Filebutton');
    const loadingFile = document.getElementById('loading-File');

    submitFileButton.addEventListener('click', function (event) {
        event.preventDefault();
        if (fileForm.checkValidity()) {
            submitFileButton.style.display = "none";
            loadingFile.style.display = 'block';
            fileForm.submit();
        } else {
            fileForm.reportValidity();
        }
    });

    const submitFolderButton = document.getElementById('submit-Folderbutton');
    const loadingFolder = document.getElementById('loading-Folder');

    submitFolderButton.addEventListener('click', function (event) {
        event.preventDefault();
        if (folderForm.checkValidity()) {
            submitFolderButton.style.display = "none";
            loadingFolder.style.display = 'block';
            folderForm.submit();
        } else {
            folderForm.reportValidity();
        }
    });
});
