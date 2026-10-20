"""Perfect-foresight solver for models with one lag and one lead (Newton on the stacked system).

residual(X) -> (T, n) array for the path X (T, n); the model supplies its own lag at t = 0 and
lead at T + 1. The Jacobian is sparse (block tridiagonal in time) and is built by finite
differences with a 3-colouring of the time index, so it costs 3n residual evaluations.
"""
import numpy as np
from scipy.sparse import csc_matrix
from scipy.sparse.linalg import spsolve


def jacobian(F, X, h=1e-7):
    T, n = X.shape
    F0 = F(X)
    rows, cols_, vals = [], [], []                 # assembled once: element-wise insertion into a
    for j in range(n):                             # lil_matrix took nearly all the run time
        for c in range(3):
            P = X.copy()
            cols = np.arange(c, T, 3)
            P[cols, j] += h
            D = (F(P) - F0) / h
            for tt in cols:
                for s in (tt - 1, tt, tt + 1):
                    if 0 <= s < T:
                        nz = np.nonzero(D[s])[0]
                        rows.append(s * n + nz)
                        cols_.append(np.full(len(nz), tt * n + j))
                        vals.append(D[s][nz])
    J = csc_matrix((np.concatenate(vals), (np.concatenate(rows), np.concatenate(cols_))), shape=(T * n, T * n))
    return J, F0


def solve_path(F, X0, tol=1e-10, maxit=40, verbose=False):
    X = X0.copy()
    for it in range(maxit):
        J, F0 = jacobian(F, X)
        err = np.abs(F0).max()
        if verbose:
            print(f"  iter {it}: max|F| = {err:.2e}")
        if err < tol:
            return X
        X = X - spsolve(J, F0.ravel()).reshape(X.shape)
    raise RuntimeError(f"Newton did not converge (max|F| = {err:.2e})")


def solve_static(f, x0, tol=1e-12, maxit=60, h=1e-7):
    """Newton for a steady state: f(x) -> residual vector of the same length."""
    x = np.asarray(x0, float).copy()
    for _ in range(maxit):
        f0 = f(x)
        if np.abs(f0).max() < tol:
            return x
        J = np.empty((len(x), len(x)))
        for j in range(len(x)):
            xp = x.copy(); xp[j] += h
            J[:, j] = (f(xp) - f0) / h
        x = x - np.linalg.solve(J, f0)
    raise RuntimeError(f"steady state did not converge (max|f| = {np.abs(f0).max():.2e})")
