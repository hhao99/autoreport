import React from "react";
import {
    Checkbox,
    FormControl, InputLabel, ListItemText, MenuItem,
    Select,
    makeStyles
} from '@material-ui/core';

const useStyle = makeStyles((theme) => ({
    layout: {
        margin: theme.spacing(5),
    },
    formControl: {
        margin: theme.spacing(3),
            minWidth: 120,
            maxWidth: 300,
    },
}));
export default ({filter, onChange})=> {
    const classes = useStyle();
    const [values,setValues] = React.useState([])

    const handleChange = (e) => {
        setValues(e.target.value)
        onChange(e)
    }

    return (
        <FormControl className={classes.formControl}>
            <InputLabel>{filter.name}</InputLabel>
            <Select
                value={values}
                onChange={handleChange}
                renderValue={(selected) => selected}
            >
                { filter.values.map( (v)=>(
                    <MenuItem key={v} value={v}>
                        <Checkbox checked={values.indexOf(v) > -1} />
                        <ListItemText>{v}</ListItemText>
                    </MenuItem>
                ))}
            </Select>
        </FormControl>
    );
}

