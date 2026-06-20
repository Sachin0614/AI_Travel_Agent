import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import logging

logger = logging.getLogger(__name__)


def create_demographic_chart(dataframe):
    """
    Create demographic pie chart showing traveler gender distribution
    """
    try:
        if dataframe is None or dataframe.empty:
            return go.Figure().add_annotation(text="No data available")
        
        # Get gender column
        gender_col = None
        for col in ['Gender', 'Traveler gender', 'Traveler Gender']:
            if col in dataframe.columns:
                gender_col = col
                break
        
        if gender_col is None:
            logger.warning("Gender column not found in dataframe")
            return go.Figure().add_annotation(text="Gender data not available")
        
        # Count genders
        gender_counts = dataframe[gender_col].value_counts()
        
        # Create pie chart
        fig = px.pie(
            values=gender_counts.values,
            names=gender_counts.index,
            title="👥 Traveler Demographics by Gender",
            color_discrete_sequence=['#FF6B35', '#004E89', '#1CAC78'],
            hole=0.3,  # Donut chart
        )
        
        fig.update_traces(textposition='inside', textinfo='percent+label')
        fig.update_layout(
            template='plotly_dark',
            font=dict(family="Arial", size=12),
            showlegend=True,
        )
        
        logger.info("Demographic chart created successfully")
        return fig
        
    except Exception as e:
        logger.error(f"Error creating demographic chart: {e}")
        return go.Figure().add_annotation(text="Error creating chart")


def create_destination_chart(dataframe):
    """
    Create horizontal bar chart showing top destinations
    """
    try:
        if dataframe is None or dataframe.empty:
            return go.Figure().add_annotation(text="No data available")
        
        if 'Destination' not in dataframe.columns:
            logger.warning("Destination column not found")
            return go.Figure().add_annotation(text="Destination data not available")
        
        # Get top 10 destinations
        top_destinations = dataframe['Destination'].value_counts().head(10)
        
        # Create horizontal bar chart
        fig = px.bar(
            x=top_destinations.values,
            y=top_destinations.index,
            orientation='h',
            title="📍 Top 10 Most Popular Destinations",
            labels={'x': 'Number of Travelers', 'y': 'Destination'},
            color=top_destinations.values,
            color_continuous_scale='Viridis',
        )
        
        fig.update_layout(
            template='plotly_dark',
            xaxis_title="Number of Travelers",
            yaxis_title="Destination",
            font=dict(family="Arial", size=12),
            showlegend=False,
            height=500,
        )
        
        fig.update_traces(hovertemplate='<b>%{y}</b><br>Travelers: %{x}<extra></extra>')
        
        logger.info("Destination chart created successfully")
        return fig
        
    except Exception as e:
        logger.error(f"Error creating destination chart: {e}")
        return go.Figure().add_annotation(text="Error creating chart")


def create_price_distribution_chart(dataframe):
    """
    Create histogram showing price distribution of accommodations
    """
    try:
        if dataframe is None or 'Cleaned_Price' not in dataframe.columns:
            return go.Figure().add_annotation(text="Price data not available")
        
        fig = px.histogram(
            dataframe,
            x='Cleaned_Price',
            nbins=30,
            title="💰 Accommodation Price Distribution",
            labels={'Cleaned_Price': 'Price (₹)'},
            color_discrete_sequence=['#1CAC78'],
        )
        
        fig.update_layout(
            template='plotly_dark',
            showlegend=False,
            xaxis_title="Price (₹)",
            yaxis_title="Frequency",
            font=dict(family="Arial", size=12),
        )
        
        logger.info("Price distribution chart created successfully")
        return fig
        
    except Exception as e:
        logger.error(f"Error creating price distribution chart: {e}")
        return go.Figure()


def create_destination_scatter(dataframe):
    """
    Create scatter plot for destinations with interactive features
    """
    try:
        if dataframe is None or dataframe.empty:
            return go.Figure()
        
        if 'Destination' not in dataframe.columns:
            return go.Figure()
        
        # Get destination counts
        dest_counts = dataframe['Destination'].value_counts().head(15)
        
        # Create scatter plot
        fig = go.Figure(data=[go.Scatter(
            x=dest_counts.index,
            y=dest_counts.values,
            mode='markers+lines',
            marker=dict(
                size=dest_counts.values / 2,
                color=dest_counts.values,
                colorscale='Viridis',
                showscale=True,
                colorbar=dict(title="Number of<br>Travelers"),
            ),
            line=dict(color='#FF6B35', width=2),
            hovertemplate='<b>%{x}</b><br>Travelers: %{y}<extra></extra>',
        )])
        
        fig.update_layout(
            title="🌍 Global Destination Popularity",
            xaxis_title="Destination",
            yaxis_title="Number of Travelers",
            template='plotly_dark',
            height=400,
            hovermode='closest',
        )
        
        fig.update_xaxes(tickangle=45)
        
        logger.info("Scatter plot created successfully")
        return fig
        
    except Exception as e:
        logger.error(f"Error creating scatter plot: {e}")
        return go.Figure()