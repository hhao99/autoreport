import React, {useState, useEffect} from 'react';

import {
    makeStyles,
    Button,
    AppBar,
    CssBaseline,
    Paper,
    Toolbar,
    Typography,
    Stepper,
    Step,
    StepLabel,

} from "@material-ui/core";

import config from "./config.json";
import ReportConfig from './report';


const useStyles = makeStyles( (theme) => ({
    appBar: {
        position: 'relative'
    },
    layout: {
        width: 'auto'
    },
    stepper: {
        width: 'auto'
    },
    Paper: {
        MarginTop: theme.spacing(3),
            marginBottom: theme.spacing(3),
            padding: theme.spacing(3)
    }
}));

export default () => {
    const classes = useStyles();
    const reports = [1,2,3]
    const [activeReport, setActiveReport ] = useState(0)


    const handleNext = (currentReport) => {
        setActiveReport(currentReport)
    }
    return (
        <React.Fragment>
            <CssBaseline />
            <header>
                <AppBar position={'absolute'} color={'primary'} className={classes.appBar}>
                    <Toolbar>
                        <Typography variant={"h6"} color={"default"}>VGC CVI Report Configuration Page</Typography>
                    </Toolbar>
                </AppBar>
            </header>
            <main>
                <Paper>
                    <Typography variant={"h3"} align='center'>
                        Report Configuration
                    </Typography>
                    <Stepper activeStep={activeReport} className={classes.stepper}>
                        { config.reports.map( r=> <Step key={r.name}><StepLabel>{r.name}</StepLabel></Step>)}
                    </Stepper>
                    <ReportConfig report={activeReport} onNext={handleNext}/>

                </Paper>
            </main>
            <footer>
                <Paper>
                    <Typography variant={"h6"} align='center'>
                        VGC CVI Report Automation Project Oct. 2020.
                    </Typography>

                </Paper>
            </footer>
        </React.Fragment>
    )
}