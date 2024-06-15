import { MicrosoftAccount } from "../entity/MicrosoftAccount";
import { AppDataSource } from "../data-source";
import { config } from "dotenv";
import axios from "axios";
import querystring from "querystring";
config();

type TokenResponse = {
    access_token: string;
    refresh_token: string;
};

export const createAuthUrl = (): string => {
    try {
        const client_id = process.env.MICROSOFT_CLIENT_ID;
        const redirect_uri = process.env.MICROSOFT_REDIRECT_URI;
        const scope = process.env.MICROSOFT_SCOPE;

        if (!client_id || !redirect_uri || !scope) {
            throw new Error("Missing environment variables");
        }

        const url = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${client_id}&response_type=code&redirect_uri=${redirect_uri}&scope=${scope}`;

        return url;

    }catch (error) {
        console.error(error);
        return null;
    }
}

export const sendAuthTokenRequest = async (code: string, is_token_expired: boolean): Promise<TokenResponse> => {
    const request_data = {
        'client_id': process.env.MICROSOFT_CLIENT_ID,
        'client_secret': process.env.MICROSOFT_SECRET_KEY,
        'scope': process.env.MICROSOFT_SCOPE,
        'grant_type': is_token_expired ? 'refresh_token' : 'authorization_code'
    };

    if (is_token_expired) {
        request_data['refresh_token'] = code;
    } else {
        request_data['code'] = code;
        request_data['redirect_uri'] = process.env.MICROSOFT_REDIRECT_URI;
    }

    try {
        const response = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',
            querystring.stringify(request_data),
            {
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                },
            });
        if (response.status !== 200) {
            throw new Error(response.data);
        }
        console.log(response.status, response.data);
        return { access_token: response.data.access_token, refresh_token: response.data.refresh_token };

    } catch (error) {
        console.error(error);
        return null;
    }
}

export const createMicrosoftAccount = async (access_token: string, refresh_token: string): Promise<MicrosoftAccount> => {
    let workbook_id = null;

    const profile_response = await axios.get('https://graph.microsoft.com/beta/me/profile',
    {
        headers: {
            'Authorization': `Bearer ${access_token}`
        }
    });

    if (profile_response.status !== 200) {
        throw new Error(profile_response.data);
    }

    const workbook_response = await axios.get("https://graph.microsoft.com/v1.0/me/drive/root/search(q='microservice_workbook')", {
        headers: {
          'Authorization': `Bearer ${access_token}`,
          'Content-Type': 'application/json'
        },
        params: {
          'select': 'id,name,file'
        }
      })
    console.log(workbook_response.data);

    if (workbook_response.status !== 200) {
        throw new Error(workbook_response.data);
    }
    
    if (workbook_response.data.value.length === 0) {
        throw new Error("Workbook not found");
    } else {
        workbook_id = workbook_response.data.value[0].id;
    }

    const microsoftAccount = new MicrosoftAccount();
    microsoftAccount.email = profile_response.data.emails[0].address;
    microsoftAccount.access_token = access_token;
    microsoftAccount.refresh_token = refresh_token;
    microsoftAccount.workbook_id = workbook_id;

    try {
        return await AppDataSource.getRepository(MicrosoftAccount).save(microsoftAccount);
    } catch (error) {
        console.error(error);
        return null;
    }
}