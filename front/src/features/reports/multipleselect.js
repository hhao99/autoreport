import React from "react";
import {
    Button,
    IconButton,
    Checkbox,
    Grid,
    FormControl, InputLabel, ListItemText, MenuItem,
    Select,
    Tooltip,
    makeStyles,
} from '@material-ui/core';
import SelectAllIcon from '@material-ui/icons/SelectAll'
import DeleteIcon from '@material-ui/icons/Delete'

const useStyle = makeStyles((theme) => ({
    layout: {
        margin: theme.spacing(5),
    },
    formControl: {
        margin: theme.spacing(3),
            minWidth: 200,
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
    const handleAddAll = (e) => {
        onChange({...e,target: { value: filter.values}})
    }
    const handleRemoveAll = (e) => {
        onChange({...e,target: { value: []}})
    }

    return (
        <FormControl className={classes.formControl}>
        <Grid container spacing={6}>
            <Grid item xs={3}>
                <Tooltip title='Add All'>
                    <IconButton variant="outlined" size='small' color="primary" onClick={handleAddAll}>
                        <SelectAllIcon />
                    </IconButton>
                </Tooltip>
                <Tooltip title='Remove All'>
                    <IconButton variant="outlined" size='small' color="primary" onClick={handleRemoveAll}>
                        <DeleteIcon />
                    </IconButton>
                </Tooltip>

            </Grid>

            <Grid item xs={6}>
                {filter.name}
                <Tooltip title={filter.selected.join(' ')}>
                    <Select
                        size='small'
                        multiple
                        value={values}
                        onChange={handleChange}
                        renderValue={(selected) => ''}
                    >
                        { filter.values.map( (v)=>(
                            <MenuItem key={v} value={v}>
                                <Checkbox checked={filter.selected.indexOf(v) > -1} />
                                <ListItemText>{v}</ListItemText>
                            </MenuItem>
                        ))}
                    </Select>
                </Tooltip>
            </Grid>

        </Grid>
        </FormControl>

    );
}

