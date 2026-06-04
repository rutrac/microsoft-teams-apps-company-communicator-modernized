// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useLocation, useNavigate, useParams } from 'react-router-dom';

// Minimal RR-v5-shape props for class components still using `match.params`.
// Lets us migrate the router without rewriting every class in the same step.
export interface IRouterProps {
    match: { params: Readonly<Record<string, string | undefined>> };
    location: ReturnType<typeof useLocation>;
    navigate: ReturnType<typeof useNavigate>;
}

export function withRouter<P extends IRouterProps>(
    Component: React.ComponentType<P>,
): React.FC<Omit<P, keyof IRouterProps>> {
    return function ComponentWithRouter(props) {
        const params = useParams();
        const location = useLocation();
        const navigate = useNavigate();
        return (
            <Component
                {...(props as P)}
                match={{ params }}
                location={location}
                navigate={navigate}
            />
        );
    };
}
