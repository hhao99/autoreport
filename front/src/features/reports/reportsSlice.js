import { createSlice, createAsyncThunk } from '@reduxjs/toolkit';
import axios from 'axios';
// import config from './config.json'


// const config_url = 'http://localhost:3002/config'
const config_url = 'http://localhost:5000/'


const loadConfig = createAsyncThunk('config/load',async ()=> {
    let result, data
    result = await axios.get(config_url+"config")
    let config = await result.data

    // update the filter

    result = await axios.get(config_url+'/get/pr_status')
    data = await result.data
    config.global.pr_state.values = data
    config.global.pr_state_2.values = data


    let filters = await Promise.all(config.global.filters.map(async (filter,index)=> {
        try {
            let result = await axios.get(config_url + '/get/' + filter.name)
            return {...filter, values: result.data }
        } catch( err) {
            console.log(err)
            return filter
        }
    }))
    let conf = {...config, global: { ...config.global, filters: filters}}
    return conf
})

const initialState = {
    config: {
        global:{
            filters:[]
        },
        slides:[]
    },
    loading: 'loading',
    error: ''
}
const reportsSlice = createSlice({
    name: 'report_config',
    initialState,
    reducers: {
        loadConfigStart: (state) => ({ ...state, loading: 'loading'}),
        loadConfigSuccess: (state) => ({...state, loading: 'idle'}),
        updateConfig: (state,action) => {
            console.log("update the config")
            console.log(action.payload)
            return {...state, config:action.payload}
        },
        postConfig: (state, action) => {

        }
    },
    extraReducers: {
        [loadConfig.pending]: (state)=> ({ ...state, loading:'loading'}),
        [loadConfig.fulfilled]: (state,action)=> ({
            ...state,
            loading: 'idle',
            config: action.payload
        }),
        [loadConfig.rejected]: (state,action)=> {
            console.log("loading the config failed")
            console.log(action.payload)
            return {
                ...state,
                loading: 'error',
                error: action.payload
            }
        }
    }
})

export const { loadConfigStart, loadConfigSuccess, updateConfig, postConfig }  = reportsSlice.actions;
export { loadConfig }
export const selectedConfig = (state) => state.reports
export default reportsSlice.reducer;