import type { NextApiRequest, NextApiResponse } from 'next'
import { NextResponse } from 'next/server'

type ResponseData = {
    message: string
}

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
    const response: ResponseData = {
        message: 'John Doe'
    };

    return res.status(200).json(response);
}