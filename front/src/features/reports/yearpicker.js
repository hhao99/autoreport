import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import {
    Slider,
    Typography,
} from '@material-ui/core';

const useStyles = makeStyles((theme) => ({
    layout: {
        width: 300,
        margin: theme.spacing(10)
    }
}));

export default function DatePickers({min,max, onChange}) {
    const classes = useStyles();
    const [value,setValue] = React.useState([min,max])
    const handleChange =(event,newValue)=> {
        setValue(newValue)
        onChange(event,newValue)
    }
    const valueText = (value) => {
        return "Year: " + value
    }

    return (
        <div className={classes.layout}>
            <Typography id="discrete-slider" gutterBottom>
                Year picker: {value[0]} - {value[1]}
            </Typography>
            <Slider
                min={min}
                max={max}
                value={value}
                valueLabelDisplay="auto"
                aria-labelledby='discrete-slider'
                getAriaValueText={valueText}
                onChange={handleChange}
                />
        </div>
    );
}
