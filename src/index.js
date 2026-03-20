/* eslint-disable no-console */
const fs = require('fs');

const actions = require('@actions/core');
const { google } = require('googleapis');

const credentials = actions.getInput('credentials', { required: true });
const parentFolderId = actions.getInput('parent_folder_id', { required: true });
const target = actions.getInput('target', { required: true });
const owner = actions.getInput('owner', { required: false });
const childFolder = actions.getInput('child_folder', { required: false });
const childFolder2 = actions.getInput('child_folder2', { required: false });
const overwrite = actions.getInput('overwrite', { required: false }) === 'true';
const convert = actions.getInput('convert', { required: false }) === 'true';
let filename = actions.getInput('name', { required: false });

const credentialsJSON = JSON.parse(Buffer.from(credentials, 'base64').toString());
const scopes = ['https://www.googleapis.com/auth/drive'];
const auth = new google.auth
    .JWT(credentialsJSON.client_email, null, credentialsJSON.private_key, scopes, owner);
const drive = google.drive({ version: 'v3', auth });

async function getOrCreateFolder(folderName, parentId) {
    const { data: { files } } = await drive.files.list({
        q: `name='${folderName}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields: 'files(id)',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
    });

    if (files.length > 1) {
        throw new Error(`More than one folder matches the name '${folderName}'`);
    }
    if (files.length === 1) {
        return files[0].id;
    }

    const { data: { id } } = await drive.files.create({
        resource: {
            name: folderName,
            mimeType: 'application/vnd.google-apps.folder',
            parents: [parentId],
        },
        fields: 'id',
        supportsAllDrives: true,
    });

    return id;
}

async function getUploadFolderId() {
    if (!childFolder) {
        return parentFolderId;
    }

    const childFolderId = await getOrCreateFolder(childFolder, parentFolderId);

    if (!childFolder2) {
        return childFolderId;
    }

    return getOrCreateFolder(childFolder2, childFolderId);
}

async function getFileId(targetFilename, folderId) {
    const { data: { files } } = await drive.files.list({
        q: `name='${targetFilename}' and '${folderId}' in parents`,
        fields: 'files(id)',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
    });

    if (files.length > 1) {
        throw new Error('More than one entry match the file name');
    }
    if (files.length === 1) {
        return files[0].id;
    }

    return null;
}

async function main() {
    const uploadFolderId = await getUploadFolderId();

    let localMimeType;
    let remoteMimeType;

    if (convert) {
        if (target.endsWith('xlsx')) {
            localMimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
            remoteMimeType = 'application/vnd.google-apps.spreadsheet';
        } else if (target.endsWith('csv')) {
            localMimeType = 'text/csv';
            remoteMimeType = 'application/vnd.google-apps.spreadsheet';
        }
    }

    if (!filename) {
        filename = target.split('/').pop();
        if (convert) {
            if (filename.endsWith('xlsx')) {
                filename = filename.slice(0, -5);
            } else if (target.endsWith('csv')) {
                filename = filename.slice(0, -4);
            }
        }
    }

    let fileId = null;

    if (overwrite) {
        fileId = await getFileId(filename, uploadFolderId);
    }

    const fileData = {
        mimeType: localMimeType,
        body: fs.createReadStream(target),
    };

    if (fileId === null) {
        if (overwrite) {
            actions.info(`File ${filename} does not exist yet. Creating it.`);
        } else {
            actions.info(`Creating file ${filename}.`);
        }
        const fileMetadata = {
            name: filename,
            parents: [uploadFolderId],
            mimeType: remoteMimeType,
        };

        drive.files.create({
            resource: fileMetadata,
            media: fileData,
            uploadType: 'multipart',
            fields: 'id',
            supportsAllDrives: true,
        });
    } else {
        actions.info(`File ${filename} already exists. Updating it.`);
        drive.files.update({
            fileId,
            media: fileData,
            supportsAllDrives: true,
        });
    }
}

main().catch((error) => actions.setFailed(error));
