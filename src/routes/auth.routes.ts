import express, { Router, Request, Response } from 'express';
import * as authControllers from '../controllers/auth.controllers';

const router: Router = express.Router();

router.get('/get_auth_url', async (req: Request, res: Response) => {
    const url = await authControllers.createAuthUrl();
    if (url) {
        res.status(200).send(url);
    } else {
        res.status(500).send('Internal server error');
    }
});

router.get('/get_auth_token', async (req: Request, res: Response) => {
    const code = req.query.code as string;
    console.log(code, req);
    if (!code) {
        res.status(400).send('Bad request');
        return;
    }
    const tokens = await authControllers.sendAuthTokenRequest(code, false);
    if (!tokens) {
        res.status(500).send('Internal server error');
        return;
    }
    const { access_token, refresh_token } = tokens;
    const account = await authControllers.createMicrosoftAccount(access_token, refresh_token);
    if (account) {
        res.status(200).send("Authenticated successfully");
    } else {
        res.status(500).send('Internal server error');
    }
});

export default router;




