import type { NextApiRequest, NextApiResponse } from 'next'
import { NextResponse } from 'next/server'

type ResponseData = {
    message: string
}

export async function GET(req: NextApiRequest) {
    const response: ResponseData = {
        message: 'John Doe'
    };

    return NextResponse.json(response);
}