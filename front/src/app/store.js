import {configureStore, getDefaultMiddleware} from '@reduxjs/toolkit';
import { createLogger } from 'redux-logger'
import reportsReducer from '../features/reports/reportsSlice';

export default configureStore({
    reducer: {
        reports: reportsReducer,
        devtools: true,
        middleware: getDefaultMiddleware().concat([createLogger()])
    }
})