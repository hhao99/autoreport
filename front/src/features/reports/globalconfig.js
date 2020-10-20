import React from 'react';
// import axios from 'axios';
// import defaultConfig from "./config.json";
import MultipleSelect from './multipleselect'
import SingleSelect from './select'

import { useDispatch, useSelector } from "react-redux";
import { updateConfig } from './reportsSlice'

import {
    makeStyles,
    Divider,
    Grid,
    Paper,
    FormControlLabel,
    Checkbox,
} from '@material-ui/core'
const useStyles = makeStyles( (theme) => ({
    Paper: {
        MarginTop: theme.spacing(3),
        marginBottom: theme.spacing(3),
        padding: theme.spacing(3)
    },
    image: {
        margin: theme.spacing(3),
        width: 120,
        height: 80
    }
}));
export default ({config})=> {
    // const config_url = 'http://localhost:5000'
    const classes = useStyles()
    const state = useSelector(state=> state.reports)
    const dispatch = useDispatch()

    const handleGlobalFilterChange = (e,index)=> {
        let filters = config.global.filters.map( (filter,i) => {
            if( index === i) {
                return {...filter, selected: e.target.value}
            }
            return filter
        })
        const conf = {...config, global: {...config.global, filters:[...filters]}}
        dispatch(updateConfig(conf))
    }
    const handlePRChange = (e,index)=> {
        let conf
        if(index === 0) {
            conf = {...config, global: {...config.global, pr_state: { ...config.global.pr_state, selected: [e.target.value]}}}
        }
        else {
            conf = {...config, global: {...config.global, pr_state_2: { ...config.global.pr_state, selected: [e.target.value]}}}
        }
        dispatch(updateConfig(conf))
    }
    const handlePageIncludeChange = (e,index)=> {

        // update the config array need to return new array based on the current array, can not mutate the current array
        let slides = config.slides.map( (slide,i) => {
            if( index ===i ) {
                let included = !slide.included
                return { ...slide, included }
            }
            return slide
        })
        let conf ={...config, slides}
        //setConfig(conf)
        dispatch(updateConfig(conf))
    }


    return (
        <React.Fragment>
            <Paper className={classes.Paper}>
                {
                    state.loading !== 'idle'  ? "loading the config " :
                        <React.Fragment>
                    <Grid container spacing={3}>
                        <Grid item xs={4}>
                            <SingleSelect filter={config.global.pr_state} onChange={e => handlePRChange(e, 0)}/>
                        </Grid>
                        <Grid item xs={4} >
                            <SingleSelect filter={config.global.pr_state_2} onChange={e => handlePRChange(e, 1)}/>
                        </Grid>
                    </Grid>
                    <Divider />
                    <Grid container spacing={3}>
                    {config.global.filters.map((filter, index) =>
                        <Grid key={filter.name} item xs={2}>
                            <MultipleSelect filter={filter} onChange={e => handleGlobalFilterChange(e, index)}/>
                        </Grid>
                    )}
                    </Grid>
                    <Divider />
                    <Grid container spacing={3}>
                    {
                        config.slides.map((slide, index) => {
                            return (
                                <Grid item xs={2} >
                                    <Grid conainter>
                                        <Grid item xs={3}>
                                            <img alt={slide.name} src={`/img/${slide.img}`} className={classes.image}/>
                                        </Grid>
                                        <Grid item xs={3}>
                                            <FormControlLabel
                                                control={
                                                    <Checkbox
                                                        checked={slide.included}
                                                        onChange={e => handlePageIncludeChange(e, index)}
                                                        name={slide.name}
                                                    />
                                                }
                                                label={slide.name}
                                            />
                                        </Grid>
                                    </Grid>
                                </Grid>
                            )
                        })
                    }
                    </Grid>
                    <Divider />
                        </React.Fragment>
                }
            </Paper>
        </React.Fragment>
    )
}