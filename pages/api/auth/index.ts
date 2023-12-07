import { NextApiRequest, NextApiResponse } from "next";
import axios from "axios";
import { readFileSync, writeFile, writeFileSync } from "fs";
import path from "path";

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
    const file = path.join(process.cwd(), 'tmp', 'refresh-token.json');
    const refreshTokenJSON = readFileSync(file, 'utf8');
    const { refresh_token } = JSON.parse(refreshTokenJSON);

    const body = {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
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

    writeFileSync(file, JSON.stringify(newRefreshToken));

    return data;
}
