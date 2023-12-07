import { NextApiRequest, NextApiResponse } from "next";
import { NextResponse } from "next/server";

import refreshTokenJSON from '../../../refresh-token.json';
import axios from "axios";
import { writeFile } from "fs";

type ResponseData = {
    access_token: string,
    refresh_token: string,
    token_type: string,
    scope: string,
    expires_in: number,
    ext_expires_in: number,
}

export default async function handler(req: NextApiRequest, res: NextApiResponse) {

    try {
        if (req.method !== 'GET') throw new Error('Method not allowed');

        const { access_token }: ResponseData = await getToken();

        const headers = {
            'Authorization': access_token,
            'Content-Type': 'application/json'
        };

        let count = 0;

        const urls = [
            'https://graph.microsoft.com/v1.0/me/drive/root',
            'https://graph.microsoft.com/v1.0/me/drive',
            'https://graph.microsoft.com/v1.0/drive/root',
            'https://graph.microsoft.com/v1.0/users',
            'https://graph.microsoft.com/v1.0/me/messages',
            'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            'https://graph.microsoft.com/v1.0/me/mailFolders',
            'https://graph.microsoft.com/v1.0/applications?$count=true',
            'https://graph.microsoft.com/v1.0/me/?$select=displayName,skills',
            'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
            'https://graph.microsoft.com/beta/me/outlook/masterCategories',
            'https://graph.microsoft.com/beta/me/messages?$select=internetMessageHeaders&$top=1',
            'https://graph.microsoft.com/v1.0/sites/root/lists',
            'https://graph.microsoft.com/v1.0/sites/root',
            'https://graph.microsoft.com/v1.0/sites/root/drives'
        ]

        const promises = urls.map(url => axios.get(url, { headers }))

        const responses = await Promise.all(promises);

        responses.forEach(item => item.status === 200 ? count++ : count);

        if (count === urls.length) {
            return res.status(200).json({ message: `All ${count} graph endpoints called successfully!🚀🚀` });
        }

    } catch (e) {

        console.log(e)
        return res.status(500).json(e);

    }
}

async function getToken() {
    const body = {
        'grant_type': 'refresh_token',
        'refresh_token': refreshTokenJSON.refresh_token,
        'client_id': process.env.CLIENT_ID,
        'client_secret': encodeURIComponent(process.env.CLIENT_SECRET!),
        'redirect_uri': 'http://localhost:53682/'
    };

    const request = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',
        body,
        {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
        });

    const data: ResponseData = request.data;

    const newRefreshToken = { 'refresh_token': data.refresh_token };

    writeFile('refresh-token.json', JSON.stringify(newRefreshToken), err => {
        if (err) {
            throw err;
        }
    });

    return data;
}