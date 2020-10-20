import React, { useState, useEffect } from 'react';
import { useSelector, useDispatch} from "react-redux";
import {
    CssBaseline,
    makeStyles,
    Button,
    FormControlLabel,
    Grid,
    Switch,
    Typography,
    Paper,
} from "@material-ui/core";
import config from './config.json'

import MultipleSelect from './multipleselect'

import YearPicker from './yearpicker'
import {current} from "@reduxjs/toolkit";

const useStyle = makeStyles( (theme) => ({
    formControl: {
        margin: theme.spacing(3),
        minWidth: 120,
    }
}))

export default ({report, onNext})=> {
    const [currentReport,setCurrentReport] = useState(report)
    const [state,setState] = useState({
        is_visible: true,
    })
    const classes = useStyle();

    useEffect( ()=> {
        console.log("loading...")
        return ()=> {
            console.log("unloading...")
        }
    },[])

    const handleBack = (e) => {
        if(currentReport > 0 ) setCurrentReport(currentReport-1)
    }
    const handleNext = (e) => {
        e.preventDefault()
        if(currentReport < config.reports.length-1) {
            setCurrentReport(currentReport+1)
            onNext(currentReport+1)
        }
        else if( currentReport === config.reports.length-1) {
            console.log("save")
            sessionStorage.setItem("config",JSON.stringify(config))
            let conf = JSON.parse(sessionStorage.getItem("config"))
            console.log(conf)

        }
    }
    const handleFilterChange = (e,index) => {
        let filter = config.reports[currentReport].filters[index]
        filter.selected = e.target.value
        console.log(config.reports[currentReport])
    }
    const handleGroupByChange = (e) => {
        let target = config.reports[currentReport].group_by
        target.selected = e.target.value
        console.log(config.reports[currentReport])
    }
    const handleStateChange = (e) => {
        setState({is_visible: !state.is_visible})
    }
    const handleYearRangeChange = (e,newValue)=> {
        console.log(newValue)
        config.reports[currentReport].year_range.selected = newValue
        console.log(config)
    }
    const handlePRChange = (e)=> {
        console.log(e.target.value)
        config.reports[currentReport].pr_round.selected=e.target.value
    }

    return (

        <div className={classes.layout}>
            <CssBaseline />
            <Paper>
                <h3>Report configuration for {config.reports[currentReport].name}</h3>
            </Paper>
            <Grid container spacing={3}>
                <Grid item xs={6} sm={3}>
                    <MultipleSelect filter={config.reports[currentReport].pr_round} onChange={e =>handlePRChange(e)} />
                </Grid>
                <Grid item xs={12} sm={6}>
                    <YearPicker
                        min={config.reports[currentReport].year_range.values[0]}
                        max={config.reports[currentReport].year_range.values[1]}
                        onChange={handleYearRangeChange}
                    ></YearPicker>
                </Grid>
                <Grid item xs={6} sm={3}>
                    Include in PPT?
                    <FormControlLabel
                        control={<Switch checked={state.is_visible} onChange={handleStateChange} name='visible'/>}
                            label='Included?'
                    >
                    </FormControlLabel>
                </Grid>

            </Grid>
            <Typography>Filters</Typography>
            <Grid container spacing={3}>
                {
                    config.reports[currentReport].filters.map( (filter,index) => (
                    <Grid item xs={6} sm={3}>
                    <MultipleSelect filter={filter} onChange={(e)=> {
                    handleFilterChange(e, index)}} />
                    </Grid>
                    ))
                }

            </Grid>

            <Typography>Group By</Typography>
            <Grid container spacing={3}>
                <Grid item xs={12} sm={6}>
                    <MultipleSelect filter={config.reports[currentReport].group_by} onChange={(e)=> {
                        handleGroupByChange(e)
                    }} />
                </Grid>
            </Grid>

            <Typography>Divided By config</Typography>
            <Grid container spacing={3}>
                <Grid item xs={12} sm={6}>
                    <MultipleSelect filter={config.reports[currentReport].compute_methods.divided_by} onChange={(e)=> {
                        handleGroupByChange(e)
                    }} />
                </Grid>
            </Grid>
            <Grid container spacing={3}>
                <Grid item xs={12} sm={6}>
                    <Button variant='contained' color='primary' disabled={currentReport === 0} onClick={handleBack}>
                        Back
                    </Button>
                </Grid>
                <Grid item xs={12} sm={6}>
                    <Button variant='contained' color='primary' onClick={handleNext}>
                        { currentReport < config.reports.length-1 ? 'Next' : 'Save'}
                    </Button>
                </Grid>
            </Grid>

        </div>
    )
}