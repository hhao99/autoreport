import React from 'react';
// import axios from 'axios';

import {
    makeStyles,
    Divider,
    Grid,
    Paper,
} from '@material-ui/core'
const useStyles = makeStyles( (theme) => ({
    Paper: {
        MarginTop: theme.spacing(3),
        marginBottom: theme.spacing(3),
        padding: theme.spacing(3)
    },
    image: {
        margin: theme.spacing(3),
        width: 180,
        height: 120
    },
    button: {
        position: "fixed",
        right: -1
    }
}));
export default ({config})=> {
    const classes = useStyles()

    return (
        <React.Fragment>
            <Paper>
                <Grid Container spacing={3}>
                    <Grid item xs={12}>
                        <Grid container>
                            <Grid item xs={3}>
                                PR State: {config.global.pr_state.selected[0]}
                            </Grid>
                            <Grid item xs={3}>
                                PR State Compared: {config.global.pr_state_2.selected[0]}
                            </Grid>
                        </Grid>
                    </Grid>
                    <Divider />
                    <Grid item xs={12}>
                        <Grid container>
                            <Grid item xs={3}>
                                <h3>Global Filters</h3>
                            </Grid>
                            {config.global.filters.map(filter =>{
                                return (
                                    <Grid item xs={3}>
                                        <Grid container spacing={3}>
                                            <Grid item xs={12}>
                                                {filter.name}
                                            </Grid>
                                            <Grid item xs={12}>
                                                {filter.selected.join(', ')}
                                            </Grid>
                                        </Grid>
                                    </Grid>
                                )
                            }
                            )}
                        </Grid>
                    </Grid>
                </Grid>
                <Divider />
                <Grid container spacing={3}>
                    <Grid item xs={12}>
                        <h3> Page config</h3>
                    </Grid>
                    <Grid item xs={12}>
                        {config.slides.map( (slide => {
                            return (
                                <Grid container spacing={3}>
                                    { !slide.included? slide.name + " not included" :
                                        <React.Fragment>
                                            <Grid item xs={3}>{slide.name}:</Grid>
                                            <Grid item xs={3}>local filter: {slide.filters.map( filter => {
                                                return (
                                                    <>
                                                    <em>{filter.name}</em>: {filter.selected.join(', ')}
                                                    </>
                                                )
                                            })}</Grid>
                                            <Grid item xs={3}>
                                                Group By: {slide.group_by.selected.join(', ')}
                                            </Grid>
                                            <Grid item xs={3}>
                                                Divided by: {slide.compute_methods.divided_by.selected.join(', ')}
                                            </Grid>
                                        </React.Fragment>
                                    }
                                </Grid>
                            )
                        }))}
                    </Grid>
                </Grid>
            </Paper>

        </React.Fragment>
    )
}